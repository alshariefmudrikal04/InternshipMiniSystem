from django.shortcuts import render
from django.http import HttpResponse
import pandas as pd
import io
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm

file_path = 'data/TAS-CASD.xlsx'
df = pd.read_excel(file_path)

def fuzzy_match(df, names_list):
    """
    Match any name in the list against FULLNAME using partial/fuzzy matching.
    e.g., searching 'joe' will match 'John Joe Doe', 'Joe Smith', etc.
    Also handles cases where individual tokens in the search query match tokens in the full name.
    """
    if not names_list:
        return df.iloc[0:0]  # empty dataframe

    mask = pd.Series([False] * len(df), index=df.index)

    for name in names_list:
        name = name.strip()
        if not name:
            continue

        # Split search name into tokens
        search_tokens = name.lower().split()

        for idx, fullname in df['FULLNAME'].items():
            if pd.isna(fullname):
                continue
            fullname_lower = str(fullname).lower()
            fullname_tokens = fullname_lower.split()

            # Check 1: direct substring match (e.g. 'joe' in 'john joe doe')
            if name.lower() in fullname_lower:
                mask[idx] = True
                continue

            # Check 2: all search tokens are substrings of any fullname token
            # e.g., searching 'jo do' matches 'John Doe' because 'jo' in 'john', 'do' in 'doe'
            all_tokens_match = all(
                any(st in ft for ft in fullname_tokens)
                for st in search_tokens
            )
            if all_tokens_match:
                mask[idx] = True

    return df[mask]


def index(request):
    data = None
    selected_column = None
    names_input = ''

    if request.method == "POST":
        names_raw = request.POST.get('names', '')
        selected_column = request.POST.get('column')
        names_input = names_raw

        # Support comma-separated OR newline-separated names
        names_list = [n.strip() for n in names_raw.replace('\n', ',').split(',') if n.strip()]

        filtered = fuzzy_match(df, names_list)

        if selected_column and selected_column in df.columns:
            cols = ['FULLNAME', selected_column]
            filtered = filtered[cols]

        # PDF export
        if request.POST.get('export') == 'pdf' and not filtered.empty:
            return export_pdf(filtered, names_list, selected_column)

        data = filtered.to_dict('records')

    return render(request, 'records/index.html', {
        'data': data,
        'columns': df.columns.tolist(),
        'selected_column': selected_column,
        'names_input': names_input,
    })


def export_pdf(filtered_df, names_list, selected_column):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=2 * cm,
        bottomMargin=1.5 * cm,
    )

    styles = getSampleStyleSheet()
    story = []

    # Title
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontSize=16,
        spaceAfter=6,
    )
    story.append(Paragraph("Internship Record Automation System", title_style))

    # Subtitle
    subtitle = f"Search results for: {', '.join(names_list)}"
    if selected_column:
        subtitle += f" | Column: {selected_column}"
    story.append(Paragraph(subtitle, styles['Normal']))
    story.append(Spacer(1, 0.5 * cm))

    # Table
    columns = list(filtered_df.columns)
    data_rows = [columns]  # header row
    for _, row in filtered_df.iterrows():
        data_rows.append([str(v) if v is not None else '' for v in row.values])

    col_count = len(columns)
    available_width = landscape(A4)[0] - 3 * cm
    col_width = available_width / col_count

    table = Table(data_rows, colWidths=[col_width] * col_count, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a3c5e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#eef2f7')]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))

    story.append(table)
    doc.build(story)

    buffer.seek(0)
    filename = "internship_records.pdf"
    response = HttpResponse(buffer, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response