import io
from datetime import timedelta
import pandas as pd
import streamlit as st

st.set_page_config(page_title='PR 15-Hour Summary', page_icon='📊', layout='wide')

st.markdown('''
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
[data-testid="stMetricValue"] {font-size: 1.55rem;}
</style>
''', unsafe_allow_html=True)

PO_COLS = ['Purchase Order Number', 'Purchase Order Date', 'Purchase Order Status', 'Vendor Name', 'QuantityOrdered']
PR_COLS = ['Receive Number', 'Receive Date', 'Vendor Name', 'PO Number', 'Quantity Received', 'CreatedTime', 'Status']


def clean_text(s):
    return s.astype('string').fillna('').str.strip()


def read_excel(uploaded_file, sheet_name):
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def normalise_inputs(po_df, pr_df, include_cancelled=False):
    missing_po = [c for c in PO_COLS if c not in po_df.columns]
    missing_pr = [c for c in PR_COLS if c not in pr_df.columns]
    if missing_po:
        raise ValueError('PO file is missing columns: ' + ', '.join(missing_po))
    if missing_pr:
        raise ValueError('PR file is missing columns: ' + ', '.join(missing_pr))

    po = po_df[PO_COLS].copy()
    pr = pr_df[PR_COLS].copy()

    for c in ['Purchase Order Number', 'Vendor Name', 'Purchase Order Status']:
        po[c] = clean_text(po[c])
    for c in ['Receive Number', 'Receive Date', 'Vendor Name', 'PO Number', 'Status']:
        if c in pr.columns:
            pr[c] = clean_text(pr[c])

    po['Purchase Order Number'] = po['Purchase Order Number'].str.upper()
    pr['PO Number'] = pr['PO Number'].str.upper()
    po['QuantityOrdered'] = pd.to_numeric(po['QuantityOrdered'], errors='coerce').fillna(0)
    pr['Quantity Received'] = pd.to_numeric(pr['Quantity Received'], errors='coerce').fillna(0)
    po['Purchase Order Date'] = pd.to_datetime(po['Purchase Order Date'], errors='coerce')
    pr['Receive Date'] = pd.to_datetime(pr['Receive Date'], errors='coerce')
    pr['CreatedTime'] = pd.to_datetime(pr['CreatedTime'], errors='coerce')

    if not include_cancelled:
        po = po[po['Purchase Order Status'].str.lower().ne('cancelled')].copy()

    # PO is ONLY a lookup source for PO Qty. It never creates report rows.
    po_qty = (
        po.groupby('Purchase Order Number', as_index=True)['QuantityOrdered']
          .sum()
          .rename('PO Qty')
    )

    return po, pr, po_qty


def build_report(po_df, pr_df, include_cancelled=False):
    po, pr, po_qty = normalise_inputs(po_df, pr_df, include_cancelled)

    valid_times = pr['CreatedTime'].dropna()
    if valid_times.empty:
        raise ValueError('No valid CreatedTime values were found in the PR file.')

    latest_time = valid_times.max()
    cutoff = latest_time - timedelta(hours=15)

    # AUTHORITATIVE LOGIC:
    # Recent PR = PR records created within latest 15 hours from the overall latest PR CreatedTime.
    # PR MTD = all older PR records in the uploaded PR dataset.
    recent = pr[pr['CreatedTime'].ge(cutoff)].copy()
    mtd = pr[pr['CreatedTime'].lt(cutoff)].copy()

    # Main report is PR-driven: only recent PR records become rows.
    recent['PO Qty'] = recent['PO Number'].map(po_qty).fillna(0)
    recent['PO Match'] = recent['PO Number'].isin(po_qty.index)
    recent['PR Qty'] = recent['Quantity Received']
    recent['Excess'] = (recent['PR Qty'] - recent['PO Qty']).clip(lower=0)
    recent['Short'] = (recent['PO Qty'] - recent['PR Qty']).clip(lower=0)
    recent['Fill Rate %'] = recent.apply(
        lambda r: (r['PR Qty'] / r['PO Qty'] * 100) if r['PO Qty'] else 0,
        axis=1,
    )

    recent = recent.sort_values(['CreatedTime', 'Receive Number'], ascending=[True, True]).reset_index(drop=True)
    recent.insert(0, 'Sl.No', range(1, len(recent) + 1))

    recent_display = recent[[
        'Sl.No', 'CreatedTime', 'Receive Number', 'PO Number', 'Vendor Name',
        'PO Qty', 'PR Qty', 'Excess', 'Short', 'Fill Rate %', 'Status'
    ]].copy()
    recent_display = recent_display.rename(columns={
        'CreatedTime': 'PR Created Time',
        'Receive Number': 'PR No',
        'PO Number': 'PO No',
    })

    # Separate MTD sheet: older PR records are retained exactly as PR records.
    mtd_display = mtd.sort_values(['CreatedTime', 'Receive Number'], ascending=[True, True]).reset_index(drop=True)
    mtd_display.insert(0, 'Sl.No', range(1, len(mtd_display) + 1))
    mtd_display = mtd_display[[
        'Sl.No', 'CreatedTime', 'Receive Number', 'PO Number', 'Vendor Name',
        'Quantity Received', 'Status'
    ]].rename(columns={
        'CreatedTime': 'PR Created Time',
        'Receive Number': 'PR No',
        'PO Number': 'PO No',
        'Quantity Received': 'PR MTD Qty',
    })

    # Optional PO-level MTD summary for quick reference. This does not drive the main report.
    mtd_summary = (
        mtd.groupby('PO Number', as_index=False)['Quantity Received']
           .sum()
           .rename(columns={'PO Number': 'PO No', 'Quantity Received': 'PR MTD Qty'})
    )
    mtd_summary['PO Qty'] = mtd_summary['PO No'].map(po_qty).fillna(0)
    mtd_summary = mtd_summary[['PO No', 'PO Qty', 'PR MTD Qty']]

    matched_recent = recent[recent['PO Match']]
    unmatched_recent = recent[~recent['PO Match']]

    meta = {
        'latest_time': latest_time,
        'cutoff': cutoff,
        'recent_rows': len(recent),
        'mtd_rows': len(mtd),
        'recent_qty': float(recent['PR Qty'].sum()),
        'mtd_qty': float(mtd['Quantity Received'].sum()),
        'po_qty_recent': float(recent['PO Qty'].sum()),
        'excess': float(recent['Excess'].sum()),
        'short': float(recent['Short'].sum()),
        'fill_rate': (float(matched_recent['PR Qty'].sum()) / float(matched_recent['PO Qty'].sum()) * 100)
            if float(matched_recent['PO Qty'].sum()) else 0.0,
        'unmatched_recent': len(unmatched_recent),
    }

    return recent_display, mtd_display, mtd_summary, meta


def excel_bytes(recent_df, mtd_df, mtd_summary, meta):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        dashboard = pd.DataFrame({
            'Metric': [
                'Latest PR CreatedTime', '15-Hour Cutoff', 'PR Qty (Latest 15h)',
                'PR MTD (Older than 15h)', 'PO Qty Lookup (Recent PRs)',
                'Excess', 'Short', 'Fill Rate', 'Recent PR Rows', 'PR MTD Rows'
            ],
            'Value': [
                meta['latest_time'].strftime('%d-%m-%Y %H:%M'),
                meta['cutoff'].strftime('%d-%m-%Y %H:%M'),
                meta['recent_qty'], meta['mtd_qty'], meta['po_qty_recent'],
                meta['excess'], meta['short'], meta['fill_rate'] / 100,
                meta['recent_rows'], meta['mtd_rows']
            ]
        })
        dashboard.to_excel(writer, sheet_name='Dashboard', index=False, startrow=1)
        recent_df.to_excel(writer, sheet_name='PR Latest 15H', index=False, startrow=1)
        mtd_df.to_excel(writer, sheet_name='PR MTD', index=False, startrow=1)
        mtd_summary.to_excel(writer, sheet_name='PR MTD by PO', index=False, startrow=1)

    buf.seek(0)
    wb = load_workbook(buf)
    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(color='FFFFFF', bold=True)
    thin = Side(style='thin', color='D9E1F2')

    for ws in wb.worksheets:
        ws.freeze_panes = 'A3'
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[2]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(bottom=thin)
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(vertical='top', wrap_text=True)

        for col_cells in ws.columns:
            letter = get_column_letter(col_cells[0].column)
            max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 34)

    for ws_name in ['PR Latest 15H', 'PR MTD by PO']:
        ws = wb[ws_name]
        # Percentage is only in the recent sheet.
        if ws_name == 'PR Latest 15H':
            for cell in ws['J'][2:]:
                cell.number_format = '0.0%'

    ws = wb['Dashboard']
    ws['B9'].number_format = '0.0%'
    ws['A1'] = 'PR 15-HOUR SUMMARY'
    ws['A1'].font = Font(size=16, bold=True)
    ws.merge_cells('A1:B1')

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def pdf_bytes(recent_df, meta):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=6*mm, leftMargin=6*mm, topMargin=7*mm, bottomMargin=7*mm)
    styles = getSampleStyleSheet()
    elements = [Paragraph('PR 15-HOUR SUMMARY', styles['Title'])]
    elements.append(Paragraph(
        f"Latest PR CreatedTime: {meta['latest_time']:%d-%m-%Y %H:%M} | "
        f"15-hour cutoff: {meta['cutoff']:%d-%m-%Y %H:%M}",
        styles['Normal']))
    elements.append(Spacer(1, 3*mm))

    metrics = [[
        'PR Qty', 'PR MTD', 'PO Qty Lookup', 'Excess', 'Short', 'Fill Rate'
    ], [
        f"{meta['recent_qty']:,.0f}", f"{meta['mtd_qty']:,.0f}", f"{meta['po_qty_recent']:,.0f}",
        f"{meta['excess']:,.0f}", f"{meta['short']:,.0f}", f"{meta['fill_rate']:.1f}%"
    ]]
    mt = Table(metrics, colWidths=[40*mm]*6)
    mt.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1F4E78')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('GRID',(0,0),(-1,-1),0.4,colors.grey),
        ('FONTSIZE',(0,0),(-1,-1),8),
        ('TOPPADDING',(0,0),(-1,-1),5), ('BOTTOMPADDING',(0,0),(-1,-1),5),
    ]))
    elements.append(mt)
    elements.append(Spacer(1, 4*mm))

    headers = list(recent_df.columns)
    rows = [headers]
    for _, r in recent_df.iterrows():
        rows.append([
            str(r['Sl.No']), str(r['PR Created Time']), str(r['PR No']), str(r['PO No']),
            str(r['Vendor Name']), f"{r['PO Qty']:,.0f}", f"{r['PR Qty']:,.0f}",
            f"{r['Excess']:,.0f}", f"{r['Short']:,.0f}", f"{r['Fill Rate %']:.1f}%", str(r['Status'])
        ])

    widths = [9*mm, 28*mm, 25*mm, 29*mm, 43*mm, 16*mm, 16*mm, 16*mm, 16*mm, 18*mm, 18*mm]
    table = Table(rows, repeatRows=1, colWidths=widths)
    table.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1F4E78')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),6.2),
        ('GRID',(0,0),(-1,-1),0.25,colors.HexColor('#B7C9D6')),
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('ALIGN',(0,0),(0,-1),'CENTER'),
        ('ALIGN',(5,1),(-2,-1),'RIGHT'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, colors.HexColor('#F5F8FA')]),
        ('LEFTPADDING',(0,0),(-1,-1),2), ('RIGHTPADDING',(0,0),(-1,-1),2),
        ('TOPPADDING',(0,0),(-1,-1),2), ('BOTTOMPADDING',(0,0),(-1,-1),2),
    ]))
    elements.append(table)
    doc.build(elements)
    return buf.getvalue()


st.title('📊 PR 15-Hour Summary')
st.caption('PR is the driver. PO is used only to look up PO Qty. The main report contains only PR records from the latest 15 hours.')

with st.sidebar:
    st.header('Daily Upload')
    po_file = st.file_uploader('PO Raw Excel', type=['xlsx', 'xls'], key='po_upload')
    pr_file = st.file_uploader('PR Raw Excel', type=['xlsx', 'xls'], key='pr_upload')
    include_cancelled = st.checkbox('Allow cancelled PO numbers for PO Qty lookup', value=False)
    st.divider()
    st.info('Cutoff = latest PR CreatedTime in the uploaded PR file minus 15 hours.')

if not po_file or not pr_file:
    st.info('Upload both PO Raw Excel and PR Raw Excel.')
    st.stop()

try:
    po_raw = read_excel(po_file, 'PurchaseOrder')
    pr_raw = read_excel(pr_file, 'PurchaseReceive')
    recent, mtd, mtd_summary, meta = build_report(po_raw, pr_raw, include_cancelled)
except Exception as e:
    st.error(f'Could not process the files: {e}')
    st.stop()

st.success(
    f"Latest PR: {meta['latest_time']:%d-%m-%Y %H:%M} | "
    f"15-hour window starts: {meta['cutoff']:%d-%m-%Y %H:%M} | "
    f"Main report rows: {meta['recent_rows']:,}"
)

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric('PR Qty · Latest 15H', f"{meta['recent_qty']:,.0f}")
c2.metric('PR MTD · Older', f"{meta['mtd_qty']:,.0f}")
c3.metric('PO Qty Lookup', f"{meta['po_qty_recent']:,.0f}")
c4.metric('Excess', f"{meta['excess']:,.0f}")
c5.metric('Short', f"{meta['short']:,.0f}")
c6.metric('Fill Rate', f"{meta['fill_rate']:.1f}%")

if meta['unmatched_recent']:
    st.warning(f"{meta['unmatched_recent']} recent PR record(s) have no matching PO number in the uploaded PO file. Their PO Qty is shown as 0.")

st.subheader('PR Latest 15 Hours — Main Report')
st.dataframe(
    recent,
    use_container_width=True,
    hide_index=True,
    column_config={
        'PO Qty': st.column_config.NumberColumn('PO Qty', format='%.0f'),
        'PR Qty': st.column_config.NumberColumn('PR Qty', format='%.0f'),
        'Excess': st.column_config.NumberColumn('Excess', format='%.0f'),
        'Short': st.column_config.NumberColumn('Short', format='%.0f'),
        'Fill Rate %': st.column_config.NumberColumn('Fill Rate %', format='%.1f%%'),
    },
)

with st.expander('PR MTD — older than the latest 15-hour window'):
    st.dataframe(mtd, use_container_width=True, hide_index=True)

with st.expander('PR MTD by PO — reference only'):
    st.dataframe(mtd_summary, use_container_width=True, hide_index=True)

st.divider()
st.subheader('Downloads')
d1, d2 = st.columns(2)
with d1:
    st.download_button(
        '⬇️ Download Excel Summary',
        data=excel_bytes(recent, mtd, mtd_summary, meta),
        file_name=f"PR_15H_Summary_{meta['latest_time']:%Y-%m-%d_%H%M}.xlsx",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )
with d2:
    st.download_button(
        '⬇️ Download PDF (Latest 15H)',
        data=pdf_bytes(recent, meta),
        file_name=f"PR_15H_Summary_{meta['latest_time']:%Y-%m-%d_%H%M}.pdf",
        mime='application/pdf',
        use_container_width=True,
    )

st.caption(
    'PR Qty = PR records created within the latest 15 hours. '
    'PR MTD = PR records created before that 15-hour cutoff. '
    'PR MTD is excluded from Excess, Short and Fill Rate. PO data is used only for PO Qty lookup.'
)
