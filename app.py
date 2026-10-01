import io
from datetime import timedelta
import pandas as pd
import streamlit as st

st.set_page_config(page_title='PO-PR Summary', page_icon='📊', layout='wide')

st.markdown('''
<style>
.block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
[data-testid="stMetricValue"] {font-size: 1.65rem;}
.small-note {font-size: 0.85rem; color: #666;}
</style>
''', unsafe_allow_html=True)

PO_COLS = [
    'Purchase Order Number', 'Purchase Order Date', 'Purchase Order Status',
    'Vendor Name', 'QuantityOrdered'
]
PR_COLS = [
    'Receive Number', 'Receive Date', 'Vendor Name', 'PO Number',
    'Quantity Received', 'CreatedTime', 'Status'
]


def clean_text(s):
    return s.astype(str).str.strip().replace({'nan': '', 'None': ''})


def read_uploaded_excel(uploaded_file, sheet_name):
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def make_summary(po_df, pr_df, include_cancelled=False):
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
    for c in ['PO Number', 'Vendor Name', 'Receive Number', 'Status']:
        pr[c] = clean_text(pr[c])

    po['QuantityOrdered'] = pd.to_numeric(po['QuantityOrdered'], errors='coerce').fillna(0)
    pr['Quantity Received'] = pd.to_numeric(pr['Quantity Received'], errors='coerce').fillna(0)
    po['Purchase Order Date'] = pd.to_datetime(po['Purchase Order Date'], errors='coerce')
    pr['CreatedTime'] = pd.to_datetime(pr['CreatedTime'], errors='coerce')
    pr['Receive Date'] = pd.to_datetime(pr['Receive Date'], errors='coerce')

    # PO is an item-level export, so aggregate quantities to one row per PO.
    if not include_cancelled:
        po = po[po['Purchase Order Status'].str.lower().ne('cancelled')]

    po['Purchase Order Number'] = po['Purchase Order Number'].str.upper()
    pr['PO Number'] = pr['PO Number'].str.upper()

    po_group = (
        po.groupby('Purchase Order Number', as_index=False)
          .agg({
              'Purchase Order Date': 'min',
              'Vendor Name': 'first',
              'QuantityOrdered': 'sum'
          })
          .rename(columns={'QuantityOrdered': 'PO Qty'})
    )

    # The latest CreatedTime in the uploaded PR data is the report reference time.
    # Current PR = CreatedTime in the last 15 hours; PR MTD = older PRs in the upload.
    valid_times = pr['CreatedTime'].dropna()
    if valid_times.empty:
        raise ValueError('No valid CreatedTime values were found in the PR file.')
    report_time = valid_times.max()
    cutoff = report_time - timedelta(hours=15)

    recent = pr[pr['CreatedTime'].ge(cutoff)].copy()
    old = pr[pr['CreatedTime'].lt(cutoff)].copy()

    recent_qty = recent.groupby('PO Number')['Quantity Received'].sum().rename('PR Qty')
    old_qty = old.groupby('PO Number')['Quantity Received'].sum().rename('PR MTD')

    recent_nos = recent.groupby('PO Number')['Receive Number'].apply(
        lambda x: ', '.join(pd.unique(x[x.ne('')].astype(str)))
    ).rename('PR No Recent')
    old_nos = old.groupby('PO Number')['Receive Number'].apply(
        lambda x: ', '.join(pd.unique(x[x.ne('')].astype(str)))
    ).rename('PR No MTD')

    # Use all PR records in the uploaded month for total fulfillment calculations.
    total_pr = (old_qty.add(recent_qty, fill_value=0)).rename('Total PR Qty')

    out = po_group.merge(old_qty, left_on='Purchase Order Number', right_index=True, how='left')
    out = out.merge(recent_qty, left_on='Purchase Order Number', right_index=True, how='left')
    out = out.merge(total_pr, left_on='Purchase Order Number', right_index=True, how='left')
    out = out.merge(recent_nos, left_on='Purchase Order Number', right_index=True, how='left')
    out = out.merge(old_nos, left_on='Purchase Order Number', right_index=True, how='left')

    for c in ['PR MTD', 'PR Qty', 'Total PR Qty']:
        out[c] = pd.to_numeric(out[c], errors='coerce').fillna(0)

    out['Excess'] = (out['Total PR Qty'] - out['PO Qty']).clip(lower=0)
    out['Short'] = (out['PO Qty'] - out['Total PR Qty']).clip(lower=0)
    out['PO FR %'] = out.apply(
        lambda r: (r['Total PR Qty'] / r['PO Qty'] * 100) if r['PO Qty'] else 0,
        axis=1
    )

    # Display order matching the PDF, with PR MTD before PR Qty.
    out = out.rename(columns={
        'Purchase Order Date': 'Date',
        'Purchase Order Number': 'PO No',
        'Vendor Name': 'Vendor Name'
    })
    out.insert(0, 'Sl.No', range(1, len(out) + 1))

    # Show PR numbers from both buckets. MTD first because it represents older activity.
    def combine_pr(a, b):
        vals = []
        for x in [a, b]:
            if pd.notna(x) and str(x).strip():
                vals.extend([v.strip() for v in str(x).split(',') if v.strip()])
        return ' & '.join(pd.unique(vals))

    out['PR No'] = [combine_pr(a, b) for a, b in zip(out['PR No MTD'], out['PR No Recent'])]
    out['Date'] = pd.to_datetime(out['Date'], errors='coerce')
    out['Date'] = out['Date'].dt.strftime('%d-%m-%y')

    display = out[[
        'Sl.No', 'Date', 'PO No', 'PR No', 'Vendor Name',
        'PO Qty', 'PR MTD', 'PR Qty', 'Excess', 'Short', 'PO FR %'
    ]].copy()

    # Keep integer-looking quantities clean.
    for c in ['PO Qty', 'PR MTD', 'PR Qty', 'Excess', 'Short']:
        display[c] = display[c].round(3)

    meta = {
        'report_time': report_time,
        'cutoff': cutoff,
        'recent_rows': len(recent),
        'old_rows': len(old),
        'recent_qty': recent['Quantity Received'].sum(),
        'old_qty': old['Quantity Received'].sum(),
        'total_po_qty': display['PO Qty'].sum(),
        'total_pr_mtd': display['PR MTD'].sum(),
        'total_pr_qty': display['PR Qty'].sum(),
        'total_excess': display['Excess'].sum(),
        'total_short': display['Short'].sum(),
        'fill_rate': (display['PR MTD'].sum() + display['PR Qty'].sum()) / display['PO Qty'].sum() * 100 if display['PO Qty'].sum() else 0,
    }
    return display, meta


def excel_download(df, meta):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        summary = pd.DataFrame({
            'Metric': ['Total POs', 'PO Qty', 'PR MTD', 'PR Qty', 'Excess Qty', 'Short Qty', 'Fill Rate'],
            'Value': [len(df), meta['total_po_qty'], meta['total_pr_mtd'], meta['total_pr_qty'],
                      meta['total_excess'], meta['total_short'], meta['fill_rate'] / 100]
        })
        summary.to_excel(writer, sheet_name='Summary', index=False, startrow=1)
        df.to_excel(writer, sheet_name='Summary', index=False, startrow=10)

    buffer.seek(0)
    wb = load_workbook(buffer)
    ws = wb['Summary']

    ws['A1'] = 'PO / PR SUMMARY'
    ws['A1'].font = Font(size=16, bold=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
    ws['A9'] = 'PO / PR Details'
    ws['A9'].font = Font(size=12, bold=True)

    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(color='FFFFFF', bold=True)
    thin = Side(style='thin', color='D9E1F2')

    for row in [2, 11]:
        for cell in ws[row]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(bottom=thin)

    for row in ws.iter_rows(min_row=12, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(vertical='top', wrap_text=True)

    # Fill-rate metric and detail percentage formatting.
    ws['B8'] = meta['fill_rate'] / 100
    ws['B8'].number_format = '0.0%'
    for cell in ws['K'][10:]:
        cell.number_format = '0.0%'

    widths = {1: 8, 2: 12, 3: 18, 4: 30, 5: 38, 6: 12, 7: 12, 8: 12, 9: 12, 10: 12, 11: 12}
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = 'A12'
    ws.auto_filter.ref = f'A11:K{ws.max_row}'

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final.getvalue()


def pdf_download(df, meta):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=7*mm, leftMargin=7*mm, topMargin=8*mm, bottomMargin=8*mm)
    styles = getSampleStyleSheet()
    elements = [Paragraph('PO / PR SUMMARY', styles['Title'])]
    elements.append(Paragraph(
        f"Report time: {meta['report_time']:%d-%m-%Y %H:%M} | Current PR window: {meta['cutoff']:%d-%m-%Y %H:%M} to {meta['report_time']:%d-%m-%Y %H:%M}",
        styles['Normal']))
    elements.append(Spacer(1, 4*mm))

    metrics = [[
        'Total POs', 'PO Qty', 'PR MTD', 'PR Qty', 'Excess Qty', 'Short Qty', 'Fill Rate'
    ], [
        f"{len(df):,}", f"{meta['total_po_qty']:,.0f}", f"{meta['total_pr_mtd']:,.0f}",
        f"{meta['total_pr_qty']:,.0f}", f"{meta['total_excess']:,.0f}", f"{meta['total_short']:,.0f}",
        f"{meta['fill_rate']:.1f}%"
    ]]
    mt = Table(metrics, colWidths=[35*mm]*7)
    mt.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1F4E78')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('GRID',(0,0),(-1,-1),0.4,colors.grey),
        ('FONTSIZE',(0,0),(-1,-1),8),
        ('BOTTOMPADDING',(0,0),(-1,-1),5),
        ('TOPPADDING',(0,0),(-1,-1),5),
    ]))
    elements.append(mt)
    elements.append(Spacer(1, 5*mm))

    headers = list(df.columns)
    rows = [headers]
    for _, r in df.iterrows():
        rows.append([
            str(r['Sl.No']), str(r['Date']), str(r['PO No']), str(r['PR No']),
            str(r['Vendor Name']), f"{r['PO Qty']:,.0f}", f"{r['PR MTD']:,.0f}",
            f"{r['PR Qty']:,.0f}", f"{r['Excess']:,.0f}", f"{r['Short']:,.0f}",
            f"{r['PO FR %']:.1f}%"
        ])

    widths = [10*mm, 18*mm, 29*mm, 35*mm, 48*mm, 17*mm, 17*mm, 17*mm, 17*mm, 17*mm, 18*mm]
    table = Table(rows, repeatRows=1, colWidths=widths)
    table.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1F4E78')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),6.5),
        ('GRID',(0,0),(-1,-1),0.25,colors.HexColor('#B7C9D6')),
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('ALIGN',(0,0),(0,-1),'CENTER'),
        ('ALIGN',(5,1),(-1,-1),'RIGHT'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, colors.HexColor('#F5F8FA')]),
        ('LEFTPADDING',(0,0),(-1,-1),2), ('RIGHTPADDING',(0,0),(-1,-1),2),
        ('TOPPADDING',(0,0),(-1,-1),2), ('BOTTOMPADDING',(0,0),(-1,-1),2),
    ]))
    elements.append(table)
    doc.build(elements)
    return buffer.getvalue()


st.title('📊 PO / PR Summary Dashboard')
st.caption('Upload the daily PO and PR raw Excel exports. The report automatically separates PR activity into PR MTD and the latest 15-hour PR window.')

with st.sidebar:
    st.header('Upload data')
    po_file = st.file_uploader('PO Raw Excel', type=['xlsx', 'xls'], key='po')
    pr_file = st.file_uploader('PR Raw Excel', type=['xlsx', 'xls'], key='pr')
    st.divider()
    include_cancelled = st.checkbox('Include cancelled POs', value=False)
    st.info('15-hour cutoff is calculated from the latest valid PR CreatedTime in the uploaded PR file.')

if not po_file or not pr_file:
    st.info('Upload both PO Raw Excel and PR Raw Excel to generate the summary.')
    st.stop()

try:
    po_raw = read_uploaded_excel(po_file, 'PurchaseOrder')
    pr_raw = read_uploaded_excel(pr_file, 'PurchaseReceive')
    summary, meta = make_summary(po_raw, pr_raw, include_cancelled=include_cancelled)
except Exception as e:
    st.error(f'Could not process the files: {e}')
    st.stop()

st.success(f"Processed {len(summary):,} POs. PR cutoff: {meta['cutoff']:%d-%m-%Y %H:%M}")

c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
c1.metric('Total POs', f"{len(summary):,}")
c2.metric('PO Qty', f"{meta['total_po_qty']:,.0f}")
c3.metric('PR MTD', f"{meta['total_pr_mtd']:,.0f}")
c4.metric('PR Qty', f"{meta['total_pr_qty']:,.0f}")
c5.metric('Excess Qty', f"{meta['total_excess']:,.0f}")
c6.metric('Short Qty', f"{meta['total_short']:,.0f}")
c7.metric('Fill Rate', f"{meta['fill_rate']:.1f}%")

st.subheader('PO / PR Details')
st.dataframe(
    summary,
    use_container_width=True,
    hide_index=True,
    column_config={
        'PO FR %': st.column_config.NumberColumn('PO FR %', format='%.1f%%'),
        'PO Qty': st.column_config.NumberColumn('PO Qty', format='%.0f'),
        'PR MTD': st.column_config.NumberColumn('PR MTD', format='%.0f'),
        'PR Qty': st.column_config.NumberColumn('PR Qty', format='%.0f'),
        'Excess': st.column_config.NumberColumn('Excess', format='%.0f'),
        'Short': st.column_config.NumberColumn('Short', format='%.0f'),
    },
)

st.divider()
st.subheader('Download Summary')
col1, col2 = st.columns(2)
with col1:
    st.download_button(
        '⬇️ Download Clean Excel Summary',
        data=excel_download(summary, meta),
        file_name=f"PO_PR_Summary_{meta['report_time']:%Y-%m-%d_%H%M}.xlsx",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )
with col2:
    st.download_button(
        '⬇️ Download Clean PDF Summary',
        data=pdf_download(summary, meta),
        file_name=f"PO_PR_Summary_{meta['report_time']:%Y-%m-%d_%H%M}.pdf",
        mime='application/pdf',
        use_container_width=True,
    )

st.caption(
    f"PR MTD includes PRs older than 15 hours within the uploaded PR data. "
    f"Current PR Qty includes PRs from {meta['cutoff']:%d-%m-%Y %H:%M} onward. "
    f"Total fulfillment = PR MTD + PR Qty."
)
