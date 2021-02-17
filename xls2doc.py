import os
from word_template_update import DocxUpdate
import pandas as pd

df = pd.read_excel(os.path.join('data','data_src.xlsx'))
df['date_start'] = df['date_start'].dt.date
df['date_end'] = df['date_end'].dt.date
df['date_print'] = df['date_print '].dt.date
list_of_dict = df.to_dict(orient='record')

for d in list_of_dict:
    doc = DocxUpdate(
        tag_dict=d,
        identifier='name',
        template_path=os.path.join('data','template_certi.docx')
    )
    doc.paragraph_update_all()
    doc.table_update(0)
    doc.doc_save()
