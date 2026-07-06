import streamlit as st
import pandas as pd
from datetime import datetime
import re
from io import BytesIO

st.set_page_config(page_title="OMT Reporting Tool", layout="wide")

st.title("📊 OMT Reporting Tool - Full Version")
st.write("Upload Excel file and generate report")

# =========================
# FILE UPLOAD
# =========================
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

# =========================
# HELPERS
# =========================
def convert_entry_date(value):
    if value is None or value == "":
        return ""

    if isinstance(value, (int, float)):
        dt = pd.to_datetime("1899-12-30") + pd.to_timedelta(value, unit="D")
        return dt.strftime("%d/%m/%Y")

    s = str(value).strip().replace("-", "/")
    parts = s.split()[0].split("/")

    if len(parts) != 3:
        return s

    d, m, y = parts
    if len(y) == 2:
        y = "20" + y

    return f"{d.zfill(2)}/{m.zfill(2)}/{y}"


def parse_date(value):
    try:
        return datetime.strptime(convert_entry_date(value), "%d/%m/%Y")
    except:
        return None


def fmt(d):
    return d.strftime("%d/%m/%Y") if d else ""


def diff_days(a, b):
    if not a or not b:
        return ""
    return (b - a).days


def clean_text(v):
    return str(v).strip()


def clean_drug(v):
    m = re.search(r"(\d+(?:\.\d+)?)\s*mg", str(v), re.I)
    return f"{m.group(1)}mg" if m else str(v)


def clean_person(v, first=False):
    s = clean_text(v)
    s = re.sub(r"\[.*?\]", "", s)
    s = re.sub(r"\S+@\S+", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.split()[0] if first else s


def get_type(drug):
    return "Weekly" if clean_text(drug) in ["8mg", "16mg", "24mg", "32mg"] else "Monthly"


# =========================
# RUN REPORT
# =========================
if uploaded_file:

    df = pd.read_excel(uploaded_file).fillna("")

    st.success(f"File loaded ✔ Rows: {len(df)}")

    start_date = st.date_input("Start Date")
    end_date = st.date_input("End Date")

    if st.button("▶ Run OMT Report"):

        data = df.to_dict("records")

        allowed_types = {"Supply To Patient", "Reverse Entry - Supply To Patient"}
        excluded_full = {"augustine fordjour", "tomasina bell"}
        excluded_first = {"augustine", "tomasina"}

        # =========================
        # CLEAN DATA
        # =========================
        cleaned = []
        for r in data:

            orig = clean_person(r.get("Prescribed By")).lower()

            r["Entry Date"] = convert_entry_date(r.get("Entry Date"))
            r["Drug"] = clean_drug(r.get("Drug"))
            r["Who"] = clean_person(r.get("Who"))
            r["Prescribed By"] = clean_person(r.get("Prescribed By"), True)

            if r.get("Entry Type") not in allowed_types:
                continue
            if orig in excluded_full:
                continue
            if clean_person(r.get("Prescribed By")).lower() in excluded_first:
                continue

            cleaned.append(r)

        # =========================
        # FILTER DATE RANGE
        # =========================
        filtered = []
        for r in cleaned:
            dt = parse_date(r.get("Entry Date"))
            if dt and start_date <= dt.date() <= end_date:
                r["_dt"] = dt
                filtered.append(r)

        st.subheader("📌 Filtered Data")
        st.dataframe(pd.DataFrame(filtered))

        # =========================
        # GROUP BY WHO
        # =========================
        by_who = {}
        for r in filtered:
            key = clean_text(r.get("Who")).lower()
            by_who.setdefault(key, []).append(r)

        seen = set()
        final_rows = []

        for r in filtered:
            key = clean_text(r.get("Who")).lower()
            if key in seen:
                continue
            seen.add(key)

            group = sorted(by_who.get(key, []), key=lambda x: x["_dt"])

            first = group[0]["_dt"] if group else None
            last = group[-1]["_dt"] if group else None

            days = diff_days(first, last)
            type_ = get_type(r.get("Drug"))

            status = "Still in program"
            if (type_ == "Weekly" and days and days > 30) or (type_ == "Monthly" and days and days > 60):
                status = "Out of Program"

            new_r = dict(r)
            new_r["Type"] = type_
            new_r["First dose"] = fmt(first)
            new_r["Last dose"] = fmt(last)
            new_r["Days between FD & LD"] = days
            new_r["Status"] = status

            final_rows.append(new_r)

        st.subheader("📊 Individual Clients Report")
        st.dataframe(pd.DataFrame(final_rows))

        # =========================
        # DOWNLOAD EXCEL
        # =========================
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame(filtered).to_excel(writer, sheet_name="Filtered Data", index=False)
            pd.DataFrame(final_rows).to_excel(writer, sheet_name="Individual Clients", index=False)

        st.download_button(
            "⬇ Download Full Report",
            data=output.getvalue(),
            file_name="OMT_Report.xlsx"
        )
