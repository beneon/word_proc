import docx
import os
import yaml
from word_template_update import DocxUpdate

with open(os.path.join('data','tag_dict.yaml'),encoding='utf8') as _yaml:
    _yaml_data = yaml.safe_load(_yaml)
    tag_dict = _yaml_data['tag_dict_test1']



doc = DocxUpdate(
    tag_dict=tag_dict,
    identifier='name',
    template_path=os.path.join('data','template_certi.docx')
)

doc.paragraph_update_all()
doc.table_update(0)
doc.doc_save()