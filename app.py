"""
Medical Benefits Comparison — v3
A clean rebuild for ingesting carrier quotes + renewal contract,
selecting which plans to present, and generating a branded Excel.
"""

import os
import io
import json
import base64
import urllib.request
import urllib.error

from flask import Flask, request, jsonify, send_file, send_from_directory
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

try:
    from pypdf import PdfReader, PdfWriter
    PYPDF_OK = True
except ImportError:
    PYPDF_OK = False

app = Flask(__name__, static_folder="static")


# ─── PDF UTILITIES ────────────────────────────────────────────────────────────
def filter_pdf_pages(pdf_bytes, page_range_str):
    if not PYPDF_OK or not page_range_str or page_range_str.strip().lower() in ("", "all"):
        return pdf_bytes
    reader = PdfReader(io.BytesIO(pdf_bytes))
    total = len(reader.pages)
    pages = set()
    for part in page_range_str.split(","):
        part = part.strip()
        if "-" in part:
            try:
                s, e = part.split("-", 1)
                pages.update(range(max(1, int(s.strip())), min(total, int(e.strip())) + 1))
            except ValueError:
                pass
        elif part.isdigit():
            p = int(part)
            if 1 <= p <= total:
                pages.add(p)
    if not pages:
        return pdf_bytes
    writer = PdfWriter()
    for p in sorted(pages):
        writer.add_page(reader.pages[p - 1])
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


# ─── EXCEL TEMPLATE WRITER ────────────────────────────────────────────────────
LABEL_MAP = {
    "Network Availability":                              "network",
    "Deductible (Employee / Family)":                   "ded",
    "Coinsurance (Member / Carrier)":                   "coinsurance",
    "MOOP (Employee / Family)":                         "moop",
    "PCP Copay":                                        "pcp",
    "Telehealth":                                       "telehealth",
    "Specialist Copay":                                 "specialist",
    "Inpatient Hospitalization":                        "inpatient",
    "Outpatient Facility":                              "op_facility",
    "Emergency Room":                                   "er",
    "Urgent Care":                                      "urgent_care",
    "Laboratory":                                       "lab",
    "X-ray":                                            "xray",
    "Imaging (CT, MRI, PET)":                          "imaging",
    "Deductible":                                       "rx_ded",
    "Generic / Preferred / Non-preferred / Specialty":  "rx_tiers",
    "Single":                                           "rate_ee",
    "Employee + Spouse":                                "rate_es",
    "Employee + Child(ren)":                            "rate_ec",
    "Family":                                           "rate_fam",
    "Rate Guarantee":                                   "notes",
}


def build_value(field_key, plan):
    if field_key == "ded":
        return f"{plan.get('ded_ind','')} / {plan.get('ded_fam','')}".strip(" /")
    if field_key == "moop":
        return f"{plan.get('moop_ind','')} / {plan.get('moop_fam','')}".strip(" /")
    if field_key == "rx_tiers":
        parts = [plan.get(k, "") for k in ("rx_generic", "rx_preferred", "rx_nonpref", "rx_specialty")]
        return " / ".join(p for p in parts if p)
    if field_key in ("rate_ee", "rate_es", "rate_ec", "rate_fam"):
        v = plan.get(field_key)
        try:
            return float(v) if v else ""
        except (TypeError, ValueError):
            return v or ""
    return plan.get(field_key, "") or ""


def find_template_anchors(ws):
    """Locate carrier_row, plan_row, and start_col by scanning for placeholders."""
    max_col = ws.max_column or 30
    start_col, carrier_row, plan_row = 8, -1, -1

    for r in range(10, 30):
        for c in range(6, min(max_col + 1, 35)):
            raw = ws.cell(r, c).value
            if raw is None:
                continue
            v = str(raw).strip().lower()
            if v in ("plan 1", "plan #1", "plan#1"):
                start_col, plan_row = c, r
                for cr in range(r - 1, max(r - 4, 0), -1):
                    if ws.cell(cr, c).value is not None:
                        carrier_row = cr
                        break
                break
            if ("carrier 1" in v or v == "carrier1") and carrier_row < 0:
                carrier_row = r
                start_col = c
        if plan_row > 0:
            break

    if carrier_row > 0 and plan_row < 0:
        plan_row = carrier_row + 1
    if plan_row < 0:
        for r in range(10, 30):
            if ws.cell(r, 2).value == "h":
                plan_row, carrier_row = r, r - 1
                break

    return start_col, carrier_row, plan_row


def safe_set(ws, row, col, value):
    """Write to a cell, unmerging slave cells first if needed."""
    if row < 1 or col < 1:
        return
    for merged in list(ws.merged_cells.ranges):
        if (merged.min_row <= row <= merged.max_row and
                merged.min_col <= col <= merged.max_col and
                not (merged.min_row == row and merged.min_col == col)):
            ws.unmerge_cells(str(merged))
            break
    ws.cell(row, col).value = value


def copy_cell_style(src, dst):
    if src.has_style:
        dst.font = src.font.copy()
        dst.fill = src.fill.copy()
        dst.border = src.border.copy()
        dst.alignment = src.alignment.copy()
        dst.number_format = src.number_format


def write_excel(template_bytes, selected_plans, client_name, renewal_data=None):
    """
    selected_plans: list of {carrier: str, plan: dict} — already filtered
    renewal_data: optional dict with rate_ee/es/ec/fam
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = next((wb[n] for n in wb.sheetnames if "medical" in n.lower()), wb.active)
    max_row = ws.max_row

    # 1. Map row labels to row indices
    row_map = {}
    for r in range(1, max_row + 1):
        v = ws.cell(r, 3).value
        if v and str(v).strip() in LABEL_MAP:
            row_map[LABEL_MAP[str(v).strip()]] = r

    # 2. Find template anchors
    start_col, carrier_row, plan_row = find_template_anchors(ws)
    if not selected_plans:
        raise ValueError("No plans selected for output.")

    # Prepend renewal as leftmost column if provided
    if renewal_data:
        renewal_plan_dict = {k: v for k, v in renewal_data.items() if k != "carrier"}
        if not renewal_plan_dict.get("plan_name"):
            renewal_plan_dict["plan_name"] = "Current Plan"
        selected_plans = [{
            "carrier": renewal_data.get("carrier") or "Current Renewal",
            "plan": renewal_plan_dict,
            "_is_renewal": True
        }] + list(selected_plans)

    # 3. Write benefit data column-by-column
    tmpl_col = start_col
    RENEWAL_FILL = PatternFill("solid", fgColor="FFF8E1")  # soft amber for renewal column
    for idx, sp in enumerate(selected_plans):
        col = start_col + idx
        plan = sp["plan"]
        is_renewal = sp.get("_is_renewal", False)

        # Copy column width
        if idx > 0:
            tmpl_dim = ws.column_dimensions[get_column_letter(tmpl_col)]
            ws.column_dimensions[get_column_letter(col)].width = tmpl_dim.width

        # Plan name
        if plan_row > 0:
            if idx > 0:
                copy_cell_style(ws.cell(plan_row, tmpl_col), ws.cell(plan_row, col))
            safe_set(ws, plan_row, col, plan.get("plan_name") or "")
            if is_renewal:
                ws.cell(plan_row, col).fill = RENEWAL_FILL

        # Benefit rows
        for field_key, row in row_map.items():
            value = build_value(field_key, plan)
            if value == "" or value is None:
                continue
            if idx > 0:
                copy_cell_style(ws.cell(row, tmpl_col), ws.cell(row, col))
            safe_set(ws, row, col, value)
            if is_renewal:
                ws.cell(row, col).fill = RENEWAL_FILL

    # 4. Write carrier names with proper merged cells per carrier group
    if carrier_row > 0:
        groups = []
        for idx, sp in enumerate(selected_plans):
            col = start_col + idx
            name = sp["carrier"]
            if not groups or groups[-1][0] != name:
                groups.append([name, col, col])
            else:
                groups[-1][2] = col

        for name, first_col, last_col in groups:
            # Unmerge any conflicting merges
            for merged in list(ws.merged_cells.ranges):
                if (merged.min_row <= carrier_row <= merged.max_row and
                        merged.max_col >= first_col and merged.min_col <= last_col):
                    ws.unmerge_cells(str(merged))

            # Copy template style to all cells in range
            tmpl_cell = ws.cell(carrier_row, tmpl_col)
            for c in range(first_col, last_col + 1):
                copy_cell_style(tmpl_cell, ws.cell(carrier_row, c))

            ws.cell(carrier_row, first_col).value = name
            if last_col > first_col:
                ws.merge_cells(start_row=carrier_row, start_column=first_col,
                              end_row=carrier_row, end_column=last_col)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── ROUTES ───────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/extract", methods=["POST"])
def extract():
    payload = request.get_json()
    api_key = payload.pop("_apiKey", None)
    page_range = payload.pop("_pageRange", "all")

    if not api_key:
        return jsonify({"error": {"message": "No API key provided."}}), 400

    # Apply page filter
    try:
        pdf_b64 = payload["messages"][0]["content"][0]["source"]["data"]
        if page_range and page_range.strip().lower() not in ("", "all"):
            pdf_bytes = base64.b64decode(pdf_b64)
            filtered = filter_pdf_pages(pdf_bytes, page_range)
            payload["messages"][0]["content"][0]["source"]["data"] = base64.b64encode(filtered).decode()
    except (KeyError, IndexError):
        pass

    try:
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=json.dumps(payload).encode(),
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=120) as resp:
            return app.response_class(response=resp.read(), status=200, mimetype="application/json")
    except urllib.error.HTTPError as e:
        return app.response_class(response=e.read(), status=e.code, mimetype="application/json")
    except Exception as e:
        return jsonify({"error": {"message": str(e)}}), 500


@app.route("/api/generate", methods=["POST"])
def generate():
    payload = request.get_json()
    template_b64 = payload.get("template")
    selected_plans = payload.get("selectedPlans", [])
    client_name = payload.get("clientName", "Group")
    renewal_data = payload.get("renewalData")

    if not template_b64:
        return jsonify({"error": {"message": "No template file provided."}}), 400

    try:
        template_bytes = base64.b64decode(template_b64)
        buf = write_excel(template_bytes, selected_plans, client_name, renewal_data)
        safe_name = client_name.replace(" ", "_") or "Group"
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"{safe_name}_Medical_Comparison.xlsx"
        )
    except Exception as e:
        return jsonify({"error": {"message": str(e)}}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8742))
    app.run(host="0.0.0.0", port=port)
