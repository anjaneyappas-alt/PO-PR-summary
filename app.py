import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================================================
# PAGE
# =========================================================

st.set_page_config(
    page_title="PO PR Summary",
    page_icon="📊",
    layout="wide"
)

st.title("📊 PO / PR Summary")

# =========================================================
# CLEANING FUNCTIONS
# =========================================================

def clean_text(series):
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
    )


def prepare_po(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    required = [
        "Purchase Order Number",
        "Purchase Order Date",
        "Purchase Order Status",
        "Vendor Name",
        "QuantityOrdered"
    ]

    missing = [x for x in required if x not in df.columns]

    if missing:
        raise Exception(
            "PO file missing columns: " + ", ".join(missing)
        )

    df["Purchase Order Number"] = clean_text(
        df["Purchase Order Number"]
    )

    df["Vendor Name"] = clean_text(
        df["Vendor Name"]
    )

    df["Purchase Order Status"] = clean_text(
        df["Purchase Order Status"]
    )

    df["Purchase Order Date"] = pd.to_datetime(
        df["Purchase Order Date"],
        errors="coerce"
    )

    df["QuantityOrdered"] = pd.to_numeric(
        df["QuantityOrdered"],
        errors="coerce"
    ).fillna(0)

    return df


def prepare_pr(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    required = [
        "Receive Number",
        "Receive Date",
        "Vendor Name",
        "PO Number",
        "Quantity Received",
        "CreatedTime",
        "Status"
    ]

    missing = [x for x in required if x not in df.columns]

    if missing:
        raise Exception(
            "PR file missing columns: " + ", ".join(missing)
        )

    df["PO Number"] = clean_text(df["PO Number"])

    df["Receive Number"] = clean_text(
        df["Receive Number"]
    )

    df["Vendor Name"] = clean_text(
        df["Vendor Name"]
    )

    df["Quantity Received"] = pd.to_numeric(
        df["Quantity Received"],
        errors="coerce"
    ).fillna(0)

    df["CreatedTime"] = pd.to_datetime(
        df["CreatedTime"],
        errors="coerce"
    )

    df["Receive Date"] = pd.to_datetime(
        df["Receive Date"],
        errors="coerce"
    )

    df["Status"] = clean_text(df["Status"])

    return df


# =========================================================
# PO SUMMARY
# =========================================================

def create_po_summary(po, include_cancelled):

    if not include_cancelled:

        status = (
            po["Purchase Order Status"]
            .str.lower()
            .str.strip()
        )

        po = po[
            ~status.str.contains(
                "cancel",
                na=False
            )
        ].copy()

    summary = (
        po.groupby(
            "Purchase Order Number",
            as_index=False
        )
        .agg({
            "Purchase Order Date": "min",
            "Vendor Name": "first",
            "QuantityOrdered": "sum"
        })
    )

    summary.rename(
        columns={
            "Purchase Order Number": "PO No",
            "Purchase Order Date": "Date",
            "QuantityOrdered": "PO Qty"
        },
        inplace=True
    )

    return summary


# =========================================================
# PR CALCULATION
# =========================================================

def calculate_pr(pr):

    """
    FOR EACH PO:

    1. Find the latest CreatedTime for that PO.
    2. Go back exactly 15 hours.
    3. PR Qty = PRs created within those 15 hours.
    4. PR MTD = PRs created before that 15-hour cutoff.

    PR MTD is NOT added to PR Qty.
    """

    rows = []

    pr = pr[
        (pr["PO Number"] != "") &
        (pr["CreatedTime"].notna())
    ].copy()

    for po_no, group in pr.groupby(
        "PO Number",
        sort=False
    ):

        group = group.sort_values(
            "CreatedTime"
        ).copy()

        # ---------------------------------------------
        # Latest PR time FOR THIS PO
        # ---------------------------------------------

        latest_time = group["CreatedTime"].max()

        # ---------------------------------------------
        # 15 HOURS BEFORE LATEST PR
        # ---------------------------------------------

        cutoff = latest_time - timedelta(hours=15)

        # ---------------------------------------------
        # PR WITHIN 15 HOURS
        # ---------------------------------------------

        recent = group[
            group["CreatedTime"] >= cutoff
        ].copy()

        # ---------------------------------------------
        # PR OLDER THAN 15 HOURS
        # ---------------------------------------------

        old = group[
            group["CreatedTime"] < cutoff
        ].copy()

        # ---------------------------------------------
        # QUANTITIES
        # ---------------------------------------------

        pr_qty = recent["Quantity Received"].sum()

        pr_mtd = old["Quantity Received"].sum()

        # ---------------------------------------------
        # RECENT PR NUMBERS
        # ---------------------------------------------

        pr_numbers = list(
            dict.fromkeys(
                recent["Receive Number"]
                .dropna()
                .astype(str)
                .str.strip()
                .tolist()
            )
        )

        rows.append({
            "PO No": po_no,
            "PR No": ", ".join(pr_numbers),
            "PR MTD": pr_mtd,
            "PR Qty": pr_qty
        })

    return pd.DataFrame(rows)


# =========================================================
# FINAL REPORT
# =========================================================

def create_report(po, pr, include_cancelled):

    po_summary = create_po_summary(
        po,
        include_cancelled
    )

    pr_summary = calculate_pr(pr)

    result = po_summary.merge(
        pr_summary,
        on="PO No",
        how="left"
    )

    result["PR No"] = result["PR No"].fillna("")

    result["PR MTD"] = result["PR MTD"].fillna(0)

    result["PR Qty"] = result["PR Qty"].fillna(0)

    # =====================================================
    # EXCESS
    # ONLY PR QTY
    # =====================================================

    result["Excess"] = (
        result["PR Qty"] -
        result["PO Qty"]
    ).clip(lower=0)

    # =====================================================
    # SHORT
    # ONLY PR QTY
    # =====================================================

    result["Short"] = (
        result["PO Qty"] -
        result["PR Qty"]
    ).clip(lower=0)

    # =====================================================
    # FILL RATE
    # ONLY PR QTY
    # =====================================================

    result["PO FR %"] = 0.0

    mask = result["PO Qty"] > 0

    result.loc[mask, "PO FR %"] = (
        result.loc[mask, "PR Qty"]
        /
        result.loc[mask, "PO Qty"]
        * 100
    )

    # =====================================================
    # SERIAL NUMBER
    # =====================================================

    result.insert(
        0,
        "Sl.No",
        range(1, len(result) + 1)
    )

    # =====================================================
    # COLUMN ORDER
    # =====================================================

    result = result[
        [
            "Sl.No",
            "Date",
            "PO No",
            "PR No",
            "Vendor Name",
            "PO Qty",
            "PR MTD",
            "PR Qty",
            "Excess",
            "Short",
            "PO FR %"
        ]
    ]

    return result


# =========================================================
# EXCEL FILE
# =========================================================

def make_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="PO PR Summary",
            index=False,
            startrow=3
        )

        ws = writer.sheets["PO PR Summary"]

        # =================================================
        # TITLE
        # =================================================

        ws["A1"] = "PO / PR SUMMARY"

        ws["A1"].font = Font(
            bold=True,
            size=18
        )

        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=11
        )

        ws["A1"].alignment = Alignment(
            horizontal="center"
        )

        # =================================================
        # DESCRIPTION
        # =================================================

        ws["A2"] = (
            "PR Qty = latest 15 hours for each PO | "
            "PR MTD = older PR quantity for the same PO"
        )

        ws.merge_cells(
            start_row=2,
            start_column=1,
            end_row=2,
            end_column=11
        )

        ws["A2"].alignment = Alignment(
            horizontal="center"
        )

        ws["A2"].font = Font(
            italic=True,
            size=10
        )

        # =================================================
        # HEADER
        # =================================================

        header_row = 4

        fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        font = Font(
            bold=True,
            color="FFFFFF"
        )

        thin = Side(
            style="thin",
            color="B7B7B7"
        )

        border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        for cell in ws[header_row]:

            cell.fill = fill
            cell.font = font
            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # =================================================
        # DATA
        # =================================================

        start = 5
        end = start + len(df) - 1

        for row in ws.iter_rows(
            min_row=start,
            max_row=end,
            min_col=1,
            max_col=11
        ):

            for cell in row:

                cell.border = border

                cell.alignment = Alignment(
                    vertical="center"
                )

        # =================================================
        # NUMBER FORMATS
        # =================================================

        for row in range(start, end + 1):

            # PO Qty / PR MTD / PR Qty /
            # Excess / Short
            for col in [6, 7, 8, 9, 10]:

                ws.cell(
                    row=row,
                    column=col
                ).number_format = '#,##0.00'

            # Fill Rate
            ws.cell(
                row=row,
                column=11
            ).number_format = '0.00"%"'

        # =================================================
        # DATE FORMAT
        # =================================================

        for row in range(start, end + 1):

            ws.cell(
                row=row,
                column=2
            ).number_format = "dd-mm-yyyy"

        # =================================================
        # CENTER NUMBERS
        # =================================================

        for row in range(start, end + 1):

            for col in [
                1, 2, 6, 7, 8, 9, 10, 11
            ]:

                ws.cell(
                    row=row,
                    column=col
                ).alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

        # =================================================
        # FREEZE
        # =================================================

        ws.freeze_panes = "A5"

        # =================================================
        # FILTER
        # =================================================

        ws.auto_filter.ref = (
            f"A4:K{end}"
        )

        # =================================================
        # TABLE
        # =================================================

        if len(df) > 0:

            table = Table(
                displayName="POPRSummary",
                ref=f"A4:K{end}"
            )

            table_style = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False
            )

            table.tableStyleInfo = table_style

            ws.add_table(table)

        # =================================================
        # COLUMN WIDTH
        # =================================================

        widths = {
            "A": 9,
            "B": 14,
            "C": 20,
            "D": 30,
            "E": 30,
            "F": 14,
            "G": 14,
            "H": 14,
            "I": 14,
            "J": 14,
            "K": 14
        }

        for col, width in widths.items():

            ws.column_dimensions[col].width = width

        ws.row_dimensions[1].height = 30
        ws.row_dimensions[4].height = 25

    output.seek(0)

    return output


# =========================================================
# UPLOAD
# =========================================================

st.subheader("Upload Raw Files")

col1, col2 = st.columns(2)

with col1:

    po_file = st.file_uploader(
        "PO Raw Excel",
        type=["xlsx", "xls"]
    )

with col2:

    pr_file = st.file_uploader(
        "PR Raw Excel",
        type=["xlsx", "xls"]
    )

include_cancelled = st.checkbox(
    "Include Cancelled POs"
)


# =========================================================
# PROCESS
# =========================================================

if po_file and pr_file:

    try:

        with st.spinner(
            "Processing..."
        ):

            po_raw = pd.read_excel(
                po_file,
                sheet_name="PurchaseOrder"
            )

            pr_raw = pd.read_excel(
                pr_file,
                sheet_name="PurchaseReceive"
            )

            po = prepare_po(po_raw)

            pr = prepare_pr(pr_raw)

            report = create_report(
                po,
                pr,
                include_cancelled
            )

        st.success(
            f"Done — {len(report):,} POs processed."
        )

        # =================================================
        # SUMMARY
        # =================================================

        total_po = len(report)

        total_po_qty = report["PO Qty"].sum()

        total_pr_mtd = report["PR MTD"].sum()

        total_pr_qty = report["PR Qty"].sum()

        total_excess = report["Excess"].sum()

        total_short = report["Short"].sum()

        fill_rate = (
            total_pr_qty /
            total_po_qty *
            100
            if total_po_qty > 0
            else 0
        )

        st.subheader("Summary")

        a, b, c, d = st.columns(4)

        a.metric(
            "Total POs",
            f"{total_po:,}"
        )

        b.metric(
            "PO Qty",
            f"{total_po_qty:,.0f}"
        )

        c.metric(
            "PR MTD",
            f"{total_pr_mtd:,.0f}"
        )

        d.metric(
            "PR Qty - 15 Hours",
            f"{total_pr_qty:,.0f}"
        )

        e, f, g = st.columns(3)

        e.metric(
            "Excess",
            f"{total_excess:,.0f}"
        )

        f.metric(
            "Short",
            f"{total_short:,.0f}"
        )

        g.metric(
            "Fill Rate",
            f"{fill_rate:.2f}%"
        )

        # =================================================
        # PREVIEW
        # =================================================

        st.subheader("Report")

        preview = report.copy()

        preview["Date"] = pd.to_datetime(
            preview["Date"],
            errors="coerce"
        ).dt.strftime("%d-%m-%Y")

        st.dataframe(
            preview,
            use_container_width=True,
            hide_index=True
        )

        # =================================================
        # DOWNLOAD
        # =================================================

        st.subheader("Download")

        excel = make_excel(report)

        st.download_button(
            "⬇️ Download Excel",
            data=excel,
            file_name="PO_PR_Summary.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    except Exception as error:

        st.error("Something went wrong.")

        st.exception(error)

else:

    st.info(
        "Upload both PO Raw Excel and PR Raw Excel files."
    )
