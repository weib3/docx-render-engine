## docx-render-engine



#### 功能

------

- 输入json，生成docx文件。



#### 依赖

---

- docx==0.8.10
- beautifulsoup==4.8.0
- Pillow==8.4.0
- flask==2.2.2
- requests==2.21.0
- pywin32



#### 快速开始

---

输入文件input.json：

```json
[
    {
        "type": "paragraph",
        "value": "XXXX医院检验科\n<r bold='False' size='10' underline='True'><r italic='True'>BRCA1/2</r>平台检测报告</r>",
        "style":
            {
                "paragraph":
                {
                    "alignment": 1
                },
                "font":
                {
                    "bold": "True",
                    "name": "微软雅黑",
                    "size": 16
                }
            }
	},
    {
        "type": "paragraph",
        "value": "",
        "style": ""
    },
    {
        "type": "paragraph",
        "value": "基本信息",
        "style":
            {
                "paragraph":
                {
                    "alignment": 0,
                    "outline_lvl": "1"
                },
                "font":
                {
                    "bold": "True",
                    "name": "微软雅黑",
                    "size": 12
                }
            }
    },
    {
    "type": "table",
    "value": [
        ["姓名", "xxx", "性别", "xx", "年龄", "xx"],
        ["住院号", "xx", "病理号", "xx", "送检科室", "xx"]
    ],
    "style": [
        {
            "selectors": ["table"],
            "style":
            {
                "alignment": 1,
                "border":
                {
                    "top":
                    {
                        "sz": 6,
                        "val": "single",
                        "color": "#000000"
                    },
                    "bottom":
                    {
                        "sz": 6,
                        "val": "single",
                        "color": "#000000"
                    },
                    "start":
                    {
                        "sz": 6,
                        "val": "single",
                        "color": "#000000"
                    },
                    "end":
                    {
                        "sz": 6,
                        "val": "single",
                        "color": "#000000"
                    }
                },
                "paragraph":
                {
                    "alignment": 1,
                    "space_before": 1,
                    "space_after": 1
                },
                "font":
                {
                    "name": "微软雅黑",
                    "size": "8"
                },
                "vertical_alignment": 1,
                "margin":
                {
                    "start": 113,
                    "top": 113,
                    "end": 113,
                    "bottom": 113
                }
            }
        }
    ]
	}
]
```

生成文档：

```python
import json
from src.document import MyDocument

f = open("demo.json", "r", encoding="UTF-8")
data = json.load(f)
f.close()

document = MyDocument()
document.render(data, "demo.docx")
```



#### 输入数据介绍

---

输入数据为list，list中每个元素为一个dict（以下简称“元素”），每个元素由3个部分定义组成：`type`，`value`，`style`。type类型介绍如下：

##### front_cover/end_cover 

生成封底或封面（注：每个封面/封底会创建新的section）。

- value：图片全路径或base64字符串，如：`/home/cover.jpg`或`base64;png;base64xxxxxxx`

- style：空

- 示例1：

```json
{
    "type": "front_cover",
    "value": "/home/cover.jpg",
    "style": ""
}
```

- 示例2：

```json
{
    "type": "front_cover",
    "value": "base64;png;base64xxxxxxx",
    "style": ""
}
```



##### paragraph

生成段落

- value: xml like文本（关于xml like文本详见xml like文本介绍）。

- style：包括2种类型：paragraph（设置整个段落段落格式）和font（设置整个段落字体格式）。详见paragraph介绍和font介绍。

- 示例：

```json
{
    "type": "paragraph",
    "value": "▶ 此患者外周血中未检测到HRR通路及遗传性肿瘤关键基因存在胚系致病或可能致病变异；",
    "style":
    {
        "paragraph":
        {
            "alignment": 3,
            "space_before": 1,
            "space_after": 1
        },
        "font":
        {
            "bold": "False",
            "name": "微软雅黑",
            "size": 10
        }
    }
}
```



##### table

生成表格。

- value：二维数组，数组中的每个元素为xml like文本（关于xml like文本详见xml like文本介绍）。单元格合并："^^"代表该单元格会向上合并；"~~"代表该单元格会向左合并。

- style：style由list组成，list中的每个元素包含selectors和style 2部分，其中：
  - `selectors`代表选择器，可以选择整个表、行、列或单元格。
    - 按正数索引：selectors: [table, row.1, row.2, column.1, column.2, cell.1.2, cell.2.3]
    - 还可以选择奇数（odd）、偶数（even）、负切片索引，如：
      - row.-1：最后一行
      - column.-2：倒数第二行
      - row.even：偶数行
      - row.odd：奇数行
      - cell.odd.even：奇数行，偶数列的单元格
  - `style`代表表格样式，具体参加表格样式。
- 示例

```json
{
    "type": "table",
    "value": [
        ["体系致病或可能致病变异", "~~", "~~", "~~", "~~"],
        ["未检测到该患者肿瘤组织样本中存在<r italic='True'>BRCA</r>基因体系致病或可能致病变异。", "~~", "~~", "~~", "]
    ],
    "style": [
    {
        "selectors": ["row.1", "cell.1.1", "column.1"],
        "style":
        {
            "shading": "#007537",
            "font":
            {
                "color": "#ffffff",
                "bold": "True",
                "size": 11
            },
            "margin":
            {
                "start": 113,
                "top": 113,
                "end": 113,
                "bottom": 113
            }
        }
    },
    {
        "selectors": ["table"],
        "style":
        {
            "alignment": 1,
            "border":
            {
                "top":
                {
                    "sz": 6,
                    "val": "single",
                    "color": "#000000"
                },
                "bottom":
                {
                    "sz": 6,
                    "val": "single",
                    "color": "#000000"
                },
                "start":
                {
                    "sz": 6,
                    "val": "single",
                    "color": "#000000"
                },
                "end":
                {
                    "sz": 6,
                    "val": "single",
                    "color": "#000000"
                }
            },
            "paragraph":
            {
                "alignment": 1,
                "space_before": 1,
                "space_after": 1
            },
            "font":
            {
                "name": "微软雅黑",
                "size": "8"
            },
            "vertical_alignment": 1,
            "margin":
            {
                "start": 113,
                "top": 113,
                "end": 113,
                "bottom": 113
            }
        }
    }]
}
```



##### header/footer

添加页眉/页脚。

- value：xml like文本（关于xml like文本详见xml like文本介绍）。

- style：style分奇数和偶数页，可以设置如下格式：
  - `odd_font`：奇数页字体；
  - `odd_paragraph`：奇数页段落格式；
  - `even_font`：偶数页字体；
  - `even_paragraph`：偶数页段落格式。
- 示例：

```json
{
    "type": "header",
    "value": "<pic src='logo.png' height='720000'></pic>\n检测报告单",
    "style":
    {
        "odd_font":
        {
            "bold": "False",
            "name": "微软雅黑",
            "size": 10
        },
        "odd_paragraph":
        {
            "alignment": 1
        },
        "even_font":
        {
            "bold": "False",
            "name": "微软雅黑",
            "size": 10
        },
        "even_paragraph":
        {
            "alignment": 1
        }
    }
}
```



##### sign

签名。

- value：value由2部分组成，sign和stamp，其中：
  - `sign`代表签名表格，二维数组，数组中的每个元素为xml like文本（关于xml like文本详见xml like文本介绍）。
  - `stamp`代表公司章。
- style：与表格格式一样。

- 示例

```json
{
    "type": "sign",
    "value": {
        "sign": [
            [
                "实验操作：<pic src='lyp.png' height='10'></pic>",
                "报告分析：<pic src='jj.png' height='10'></pic>",
                "报告初审：<pic src='lc.png' height='10'></pic>",
                "报告复审：<pic src='cx.png' height='10'></pic>",
                "报告日期：2022-05-16"
            ]
        ],
        "stamp": "zhang_bj.png"
    },
    "style": [
        {
            "selectors": [
                "table"
            ],
            "style": {
                "border": {
                    "top": {
                        "sz": 6,
                        "val": "single",
                        "color": "#000000"
                    },
                    "bottom": {
                        "sz": 6,
                        "val": "single",
                        "color": "#000000"
                    },
                    "start": {
                        "sz": 6,
                        "val": "single",
                        "color": "#000000"
                    },
                    "end": {
                        "sz": 6,
                        "val": "single",
                        "color": "#000000"
                    }
                },
                "paragraph": {
                    "alignment": "1",
                    "space_after": "0"
                },
                "font": {
                    "name": "宋体",
                    "size": 8
                },
                "vertical_alignment": "1"
            }
        },
        {
            "selectors": [
                "cell.1.5"
            ],
            "style": {
                "width": "1162800"
            }
        }
    ]
}
```



##### page_break

分页。

- value：空

- style：空

- 示例

```
{
    "type": "page_break",
    "value": "",
    "style": ""
}
```



##### section

添加新节。

- value：空

- style：空

- 示例

```json
{
    "type": "section",
    "value": "",
    "style": ""
}
```



##### table_of_contents

添加目录。

- value：通过域实现，可以设置目录显示级别、目录格式等。

- style：空

- 示例

```json
{
    "type": "table_of_contents",
    "value": "TOC \\o \"1-4\" \\h \\z \\u",
    "style": ""
}
```



#### xml like文本介绍

---

xml like文本指一段可以包含或不包含标签的文本。如果包含标签，必须包含起始和终止标签。另外如果包含`\n`则会分成多个段落。

目前可以识别的标签包括：

- `<p></p>`：段落，插入段落并设置段落格式。如：

  ```xml
  <p alignemnt='1'>乳腺癌易感基因是重要的抑癌基因及肿瘤易感基因</p>
  ```

- `<r></r>`：run，插入文本并设置字体样式。如：

  ```xml
  乳腺癌易感基因是重要的抑癌基因及肿瘤易感基因<r superscript='True'>[1]</r>。
  ```

- `<pic></pic>`：图，插入图。如：

  ```xml
  <pic src='h3_icon.jpg' width='133200' height='118800'></pic>致病或可能致病的胚系变异
  ```

- `<page></page>`：设置当前页码。如：

  ```xml
  第<page></page>页
  ```

- `<sectionpages></sectionpages>`：本节总页数。如：

  ```xml
  第<page></page>页，共<sectionpages></sectionpages>页
  ```

- `<numpages></numpages>`：文档总页数。如：

  ```xml
  第<page></page>页，共<numpages></numpages>页
  ```

  

#### 段落格式

---

详细介绍见https://python-docx.readthedocs.io/en/latest/api/text.html#paragraphformat-objects

- `alignment`：对齐方式，值包括：

  - 0：左对齐
  - 1：居中对齐
  - 2：右对齐
  - 3：两端对齐
  - 4：分散对齐

- `first_line_indent`：首行缩进长度。

- `keep_together`：段落是否应该保持完整不跨页。True代表不跨页，反之为False。

- `keep_with_next`：段落是否应该与后续段落保持在同一页上。True代表与后续段落保持在同一页，反之为False。

- `line_spacing`：行间距。浮点数表示行距是行高的倍数，如果是整数代表是units值。

- `page_break_before`：该段落是否出现在前一段之后的页面顶部。True代表出现在顶部，反之为False。

- `space_after`：与下一段的间距，units值。

- `space_before`：与上一段的间距，units值。

- `widow_control`：分页时，段落第一行到最后一行是否保持在同一页上。True代表保持在同一页，反之为False。

- `outline_lvl`：大纲视图级别，取值为1，2，3，4，5等。用于在导航窗格显示级别。

- `border`：段落边框，取值参考：http://officeopenxml.com/WPtableBorders.php

  - top：代表上边框
  - bottom：代表下边框
  - start：代表左边框
  - end：代表有边框

  

#### 字体格式

---

详细介绍见https://python-docx.readthedocs.io/en/latest/api/text.html#font-objects

- `bold`：字体加粗，True为加粗，反之为False。

- `color`：字体颜色，HEX值。如：#000000。

- `italic`：字体倾斜，True为倾斜，反之为False。

- `name`：字体名称。如：宋体。

- `shadow`：字体阴影，True为有阴影，反之为False。

- `size`：字体大小，units值。

- `subscript`：字体下标，True为下标。

- `superscript`：字体上标，True为上标。

- `underline`：下划线，True为有下划线。

  

#### Inline图格式

---

- `src`：图片全路径。
- `height`：图高度，units值。
- `width`：图宽度，units值。



#### 表格格式

---

- row/column/cell/table均可识别格式：
  - `paragraph`：段落格式
  - `font`：字体格式。
  - `border`：表格边框，取值参考：http://officeopenxml.com/WPtableBorders.php
  - `shading`：背景，HEX值。如：#000000。
  - `vertical_alignment`：段落垂直对齐方式。
    - 0：顶端对齐
    - 1：居中对齐
    - 3：底部对齐
  - `margin`：段落与边框距离。
    - top：段落与上边框距离
    - bottom：段落与下边框距离
    - start：段落与左边框距离
    - end：段落与有边框距离
- row其他格式
  - `height`：行高，units值。
- column其他格式
  - `width`：列宽，units值。
- table其他格式
  - `autofit`：表格列宽是否随内容自动调整。True为自动调整。
  - `alignment`：表格对齐方式。取值与段落的alignment格式一致。
  - `width`：表格宽度，units值。
  - `first_row_repeat_in_each_page`：表格如果跨页，表头是否重复。True为跨页表头重复。