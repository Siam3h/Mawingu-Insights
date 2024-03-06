import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors


try:
    df = pd.read_excel('SlumPublications.xlsx')
except UnicodeDecodeError:
    df = pd.read_excel('SlumPublications.xlsx', encoding='latin1')  # or 'ISO-8859-1'


df['DESCRIPTIVE TERMS/ADJECTIVES/ADJECTIVAL PHRASES (e.g. filthy, dirty, undesirable)'] = df['DESCRIPTIVE TERMS/ADJECTIVES/ADJECTIVAL PHRASES (e.g. filthy, dirty, undesirable)'].astype(str)
df['DESCRIPTIVE TERMS/ADJECTIVES/ADJECTIVAL PHRASES (e.g. filthy, dirty, undesirable)'] = df['DESCRIPTIVE TERMS/ADJECTIVES/ADJECTIVAL PHRASES (e.g. filthy, dirty, undesirable)'].fillna('')  # Replace NaNs with empty strings


text = ' '.join(df['DESCRIPTIVE TERMS/ADJECTIVES/ADJECTIVAL PHRASES (e.g. filthy, dirty, undesirable)'])  # Use the correct column name


word_freq = pd.Series(text.split()).value_counts().reset_index()
word_freq.columns = ['Word', 'Frequency']


doc = SimpleDocTemplate("word_cloud_with_frequency.pdf", pagesize=letter)
elements = []


table_data = [word_freq.columns.tolist()] + word_freq.values.tolist()


table_style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.gray),
                          ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                          ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                          ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                          ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                          ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                          ('GRID', (0, 0), (-1, -1), 1, colors.black)])

table = Table(table_data)
table.setStyle(table_style)

elements.append(table)

doc.build(elements)
