# 使用python-docx更新docx模板文档

## 为什么要写这个程序

我经常遇到要根据外部数据动态更新word文档的需求，其实word自己有邮件合并就是做这个的，但是我就是想用python来写一个工具来干同样的事情呀，况且我在word里面打几个字就能完成的事情，为什么非要用鼠标点来点去才行呢？

## 程序所需环境以及程序文件结构

程序需要python-docx库，示范程序`xls2doc.py`以及`update_template.py`分别用了pandas库读取excel文件，以及pyyaml库读取yaml文件，虽然不是必须，但是也建议安装。模板docx和数据都放在data文件夹中，使用程序输出的文档保存在output文件夹里。

可以参考上面提到的两个示例python程序（`xls2doc.py`以及`update_template.py`)了解本程序的使用

程序文件`word_template_update.py`包含了两个类，`DocxUpdate` 这个类是负责更新文档的。

```python

class DocxUpdate:
    def __init__(self,tag_dict:dict,identifier:str,template_path:str, dunder_included:bool=False):
        """

        :param tag_dict: 提供数据的字典
        :param identifier: 字典中对该记录有辨识作用的key，其value将作为输出文档的文件名
        :param template_path: 模板docx文档所在位置
        :param dunder_included: 字典中各个key字符是否包含前后两个__，默认为False
        """
```

实例化以后，常规的对各个段落的内容进行更新是使用`paragraph_update_all()`方法，而更新表格的内容则需要提供表格的index，python-docx会将文档中所有的表格从前到后放到tables这个list里面，`table_update(table_id)`这个方法会对`doc.tables[table_id]`这一个表格里面所有单元格内的paragraphs做一个文字的替换

## 难道不能直接对paragraph.text做文字的replace么？

当然是可以的，只不过这样会导致段落原有的runs全部丢失，转而新建单独的一个run。这样一来原有的格式就丢了

所以肯定要找到模板标识符(`__xxx__`)所在的run，然后去进行text的替换。原先我以为一个标识符如果格式上面没有做什么特别的设置，应该就在一个run里面的。但其实不是这样，word的奇怪机制可能会将一个标识符拆成三个部分或者更多的部分。所以程序里另外写了一个类 `ParagraphRunProc` 来将同属于一个标识符的run作合并。