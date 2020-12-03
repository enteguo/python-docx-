# Main

本文档是python-docx文档的部分翻译稿件，

目前包括四个方面，document文档，text文本，section小节，Headers and Footers页眉和页脚。

[TOC]

------

# Document文档

python-docx允许您创建新文档以及对现有文档进行更改。实际上，它仅允许您对现有文档进行更改。
只是如果您从一个没有任何内容的文档开始，乍一看就像是从头开始创建一个文档。

文档的外观很大程度上取决于删除所有内容时剩下的部分。诸如样式（styles），页眉和页脚（page headers and footers）之类与主要内容分开，在文档开始时进行大量自定义，然后应用在所生成的文档中。

让我们逐步完成创建一个文档的步骤，首先从您可以对文档执行的两项主要操作开始，将其打开并保存。

## 打开文档

最简单的入门方法是在不指定文件的情况下打开新文档：

```python
from docx import Document

document = Document()
document.save('test.docx')
```

这将从内置的默认模板创建一个新文档，并将其保存为“ test.docx”文件。所谓的“默认模板”实际上只是一个没有内容的Word文件，该文件与安装的python-docx软件包一起存储。
与选择Word中的 **File > New from Template**菜单项后选择Word文档模板所获得的效果大致相同。

## 打开一个存在的文档

如果要对文档进行更多控制，或者要更改现有文档，则需要使用文件名打开一个文档：

```python
document = Document('existing-document-file.docx')
document.save('new-file-name.docx')
```

注意事项：

- 您可以通过这种方式打开任何Word 2007或更高版本的文件（来自Word 2003及更早版本的.doc文件将不起作用）。尽管您可能还不能控制所有内容，但是其中已经存在的任何内容都可以加载和保存。该功能集仍在构建中，因此您尚不能添加或更改标题或脚注之类的内容，但是如果文档中包含这些内容，python-docx可以避免修改并保存它们。
- 如果您使用相同的文件名打开并保存文件，则python-docx将覆盖原始文件。

## 打开“类似文件”的文档

python-docx可以从所谓的类似文件的对象中打开文档。它还可以保存到类似文件的对象。
当您想通过网络或数据库获取目标文档，并且不想（或不允许）与文件系统进行交互时，这会很方便。实际上，这意味着您可以传递打开文件或StringIO/BytesIO流对象来打开或保存文档，如下所示：

```python
f = open('foobar.docx', 'rb')
document = Document(f)
f.close()

# or

with open('foobar.docx', 'rb') as f:
    source_stream = StringIO(f.read())
document = Document(source_stream)
source_stream.close()
...
target_stream = StringIO()
document.save(target_stream)
```

并非在所有操作系统上都需要'rb'文件打开模式参数。它默认为'r'，这有时足够了，但是Windows和某些版本的Linux上要求使用'b'（选择二进制模式）以允许打开Zipfile文件。

------

# TEXT（文本）

为了有效处理文本，重要的是要先对段落等block-level元素和run等inline-level对象有所了解。

## Block-level vs. inline text objects

paragraph（段落）是Word中的主要 block-level对象。block-level对象是处于左右边界之间的文本，每当文本超出其右边界时，就会增加一行。对于paragraph（段落），边界通常是页边距；但是如果页面按列布置，边界也可以是列边界；如果段落出现在表格单元格内，则边界也可以是单元格边界。

Table（表）也是block-level 对象。

Inline-level对象是出现在block-level项内的内容的一部分。举一个例子：一个以粗体显示的单词或一个全大写的句子。最常见的内联对象是run。Block-level容器中的所有内容都在inline-level对象内部。通常，一个paragraph（段落）包含一个或多个run，每个run包含该paragraph文本的一部分。

Block-level项的属性指定其在页面上的位置，比如段落前后的缩进和空格。Inline-level项的属性通常指定显示内容的外观，例如字体，字体大小，粗体和斜体。

## 段落属性

paragraph（段落）具有多种属性，可以指定其在容器（通常是page（页面））中的位置以及将内容划分为不同行的方式。

通常，最好定义一个paragraph样式，将这些属性收集到一个有意义的group中，然后将适当的样式应用于每个paragraph，而不是重复地将这些属性直接应用于每个paragraph。这类似于CSS与HTML一起工作的方式。此处描述的所有paragraph属性都可以使用style（样式）设置，也可以直接应用于paragraph。

paragraph的格式属性可使用段落的**ParagraphFormat** 对象的**paragraph_format**属性进行访问。

### 水平对齐

也称为justification（对齐），可以使用枚举类型**WD_PARAGRAPH_ALIGNMENT**将段落的对齐方式设置为左对齐，居中对齐，右对齐或完全对齐（左右对齐）：

```python
>>> from docx.enum.text import WD_ALIGN_PARAGRAPH
>>> document = Document()
>>> paragraph = document.add_paragraph()
>>> paragraph_format = paragraph.paragraph_format

>>> paragraph_format.alignment
None  # indicating alignment is inherited from the style hierarchy
>>> paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
>>> paragraph_format.alignment
CENTER (1)
```

### 缩进

Indentation（缩进）是段落和其容器边缘之间的水平空间，通常是页边距。paragraph（段落）可以在左侧和右侧分别缩进。第一行的缩进也可以与该段落的其余部分不同。*first line indent*（首行缩进）比段落的其余部分更加缩进。缩进较少的第一行具有*hanging indent*（悬挂缩进）。

缩进使用[Length](https://python-docx.readthedocs.io/en/latest/api/shared.html#docx.shared.Length)(长度)值指定，例如“[Inches](https://python-docx.readthedocs.io/en/latest/api/shared.html#docx.shared.Inches)”，“ [Pt](https://python-docx.readthedocs.io/en/latest/api/shared.html#docx.shared.Pt)”或“[Cm](https://python-docx.readthedocs.io/en/latest/api/shared.html#docx.shared.Cm)”。负值同样是有效的，并且使段落与边距重叠指定的单位数量。**None**表示缩进值从style（样式）层次结构继承。为缩进属性赋值None会删除任何直接应用的缩进设置，并从style（样式）层次结构继承：

```python
>>> from docx.shared import Inches
>>> paragraph = document.add_paragraph()
>>> paragraph_format = paragraph.paragraph_format

>>> paragraph_format.left_indent
None  # indicating indentation is inherited from the style hierarchy
>>> paragraph_format.left_indent = Inches(0.5)
>>> paragraph_format.left_indent
457200
>>> paragraph_format.left_indent.inches
0.5
```

类似的，右侧缩进：

```python
>>> paragraph_format.right_indent
None
>>> paragraph_format.right_indent = Pt(24)
>>> paragraph_format.right_indent
304800
>>> paragraph_format.right_indent.pt
24.0
```

首行缩进使用**first_line_indent**属性指定，是相对于左侧缩进进行缩进。负值表示悬挂缩进：

```python
>>> paragraph_format.first_line_indent
None
>>> paragraph_format.first_line_indent = Inches(-0.25)
>>> paragraph_format.first_line_indent
-228600
>>> paragraph_format.first_line_indent.inches
-0.25
```

### 制表位

Tab stops（制表位）决定了段落文本中制表符的呈现方式。特别是，它指定了制表符之后的文本的开始位置，该位置的对齐方式，以及一个可选的前导字符，它将填充制表符跨越的水平空间。

paragraph（段落）或style（样式）的制表位包含在 [TabStops](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.tabstops.TabStops)对象中，该对象使用[ParagraphFormat](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat)上的[tab_stops](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.tab_stops)属性访问：

```python
>>> tab_stops = paragraph_format.tab_stops
>>> tab_stops
<docx.text.tabstops.TabStops object at 0x106b802d8>
```

使用 [add_tab_stop()](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.tabstops.TabStops.add_tab_stop)方法添加一个新的制表位：

```python
>>> tab_stop = tab_stops.add_tab_stop(Inches(1.5))
>>> tab_stop.position
1371600
>>> tab_stop.position.inches
1.5
```

对齐方式默认为左对齐，但可以通过[WD_TAB_ALIGNMENT](https://python-docx.readthedocs.io/en/latest/api/enum/WdTabAlignment.html#wdtabalignment)枚举的成员来指定。前导字符默认为空格，但可以通过 [WD_TAB_LEADER](https://python-docx.readthedocs.io/en/latest/api/enum/WdTabLeader.html#wdtableader)枚举的成员来指定：

```python
>>> from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
>>> tab_stop = tab_stops.add_tab_stop(Inches(1.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
>>> print(tab_stop.alignment)
RIGHT (2)
>>> print(tab_stop.leader)
DOTS (1)
```

使用通过访问[TabStops](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.tabstops.TabStops)上的序列访问现有的制表位：

```python
>>> tab_stops[0]
<docx.text.tabstops.TabStop object at 0x1105427e8>
```

[TabStops](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.tabstops.TabStops)和[TabStop](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.tabstops.TabStop)API文档中提供了更多详细信息。

### 段落间距

 [space_before](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.space_before)和[space_after](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.space_after)属性分别控制paragraph（段落）前后的间距。在页面布局期间，段落之间的间距是重叠的，这意味着两个段落之间的间距是第一段的*space_after*和第二段的*space_before*的最大值。段落间距通常使用[Pt](https://python-docx.readthedocs.io/en/latest/api/shared.html#docx.shared.Pt)作为 [Length](https://python-docx.readthedocs.io/en/latest/api/shared.html#docx.shared.Length) 值：

```python
>>> paragraph_format.space_before, paragraph_format.space_after
(None, None)  # inherited by default

>>> paragraph_format.space_before = Pt(18)
>>> paragraph_format.space_before.pt
18.0

>>> paragraph_format.space_after = Pt(12)
>>> paragraph_format.space_after.pt
12.0
```

### 行间距

Line spacing（行间距）是一段段落中基线之间的距离。行距可以指定为绝对距离或相对于行高（基本上是所用字体的磅值）。

典型的绝对度量值为18 points。

典型的相对度量是双倍行间距（2.0行高）。默认行距是单行距（1.0行高）。

行间距由[line_spacing](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.line_spacing)和 [line_spacing_rule](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.line_spacing_rule)属性的相互作用控制。 
[`line_spacing`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.line_spacing)可以是[`Length`](https://python-docx.readthedocs.io/en/latest/api/shared.html#docx.shared.Length) 值，小数值[`float`](https://docs.python.org/3/library/functions.html#float)或None。
[`Length`](https://python-docx.readthedocs.io/en/latest/api/shared.html#docx.shared.Length) 值表示绝对距离。[`float`](https://docs.python.org/3/library/functions.html#float)表示行高的数量。None表示行距是继承的。
[`line_spacing_rule`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.line_spacing_rule)是[WD_LINE_SPACING](https://python-docx.readthedocs.io/en/latest/api/enum/WdLineSpacing.html#wdlinespacing)枚举的成员，或者是None：

```python
>>> from docx.shared import Length
>>> paragraph_format.line_spacing
None
>>> paragraph_format.line_spacing_rule
None

>>> paragraph_format.line_spacing = Pt(18)
>>> isinstance(paragraph_format.line_spacing, Length)
True
>>> paragraph_format.line_spacing.pt
18.0
>>> paragraph_format.line_spacing_rule
EXACTLY (4)

>>> paragraph_format.line_spacing = 1.75
>>> paragraph_format.line_spacing
1.75
>>> paragraph_format.line_spacing_rule
MULTIPLE (5)
```

### 分页属性

四个段落属性， [`keep_together`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.keep_together), [`keep_with_next`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.keep_with_next), [`page_break_before`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.page_break_before), 和[`widow_control`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.widow_control)控制着段落在页面边界附近的布局。

[`keep_together`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.keep_together)控制整个段落出现在同一页面上，如果该段落会在两个页面之间被分页，则会在该段落之前发出分页符。

[`keep_with_next`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.keep_with_next)将一个段落与后续段落保持在同一页面上。例如，这可用于使section heading（节标题）与section （节）的第一段在同一页面上。

[`page_break_before`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.page_break_before)将段落放置在新页面的顶部。可以在chapter heading（章节标题）上使用它，以确保chapter （章节）在新页面上开始。

[`widow_control`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.parfmt.ParagraphFormat.widow_control)分页以避免将段落的第一行或最后一行与段落的其余部分放在单独的页面上。

所有这四个属性都是*tri-state*（三态的），这意味着它们可以采用值**True**，**False**或**None**。 
**None**表示属性值是从样式层次结构继承的。
**True**表示“打开”，**False**表示“关闭”：

```python
>>> paragraph_format.keep_together
None  # all four inherit by default
>>> paragraph_format.keep_with_next = True
>>> paragraph_format.keep_with_next
True
>>> paragraph_format.page_break_before = False
>>> paragraph_format.page_break_before
False
```

## 应用字符格式

Character formatting（字符格式）在run级别应用。包括typeface（字体），size（大小），bold（粗体），italic（斜体）和underline（下划线）。

[`Run`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Run)对象具有只读字体属性，提供对[`font`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Run.font)（字体）对象的访问。
[`Run`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Run)的[`font`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Run.font)对象提供用于获取和设置[`Run`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Run)的字符格式的属性。

下面提供了几个示例。有关可用属性的完整集合，请参见[`Font`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Font) API文档。

可以通过以下方式访问run的font字体：

```python
>>> from docx import Document
>>> document = Document()
>>> run = document.add_paragraph().add_run()
>>> font = run.font
```

字体和大小设置如下：

```python
>>> from docx.shared import Pt
>>> font.name = 'Calibri'
>>> font.size = Pt(12)
```

许多字体属性是*tri-state*(三态的)，这意味着它们可以采用值**True**，**False**和**None**。 **True**表示属性为开启，**False**表示属性为关闭。从概念上讲，**None**值表示“继承”。默认情况下，run从样式继承层次结构结构继承其字符格式。可以对 [`Font`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Font)对象直接赋值任何字符格式都会覆盖继承的值。

Bold （粗体），italic （斜体），all-caps（全大写），strikethrough（删除线），superscript（上标）和其他许多属性也是如此。有关完整列表，请参见 [`Font`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Font) API文档。

```python
>>> font.bold, font.italic
(None, None)
>>> font.italic = True
>>> font.italic
True
>>> font.italic = False
>>> font.italic
False
>>> font.italic = None
>>> font.italic
None
```

Underline（下划线）有点特殊。它是tri-state （三态属性）和enumerated（枚举属性）的混合。 **True**表示单下划线，是迄今为止最常见的下划线。 
**False**表示没有下划线，但更常见的是，如果不需要下划线，则**None**是正确的选择。
用 [WD_UNDERLINE](https://python-docx.readthedocs.io/en/latest/api/enum/WdUnderline.html#wdunderline)枚举的成员指定其他形式的下划线，例如双划线或虚线：

```python
>>> font.underline
None
>>> font.underline = True
>>> # or perhaps
>>> font.underline = WD_UNDERLINE.DOT_DASH
```

### 字体颜色

每个 [`Font`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Font)对象都有一个 [`ColorFormat`](https://python-docx.readthedocs.io/en/latest/api/dml.html#docx.dml.color.ColorFormat)对象，该对象可通过其只读的 [`color`](https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.run.Font.color)属性访问其颜色。

将特定的RGB颜色应用于字体：

```python
>>> from docx.shared import RGBColor
>>> font.color.rgb = RGBColor(0x42, 0x24, 0xE9)
```

还可以通过赋值[MSO_THEME_COLOR_INDEX](https://python-docx.readthedocs.io/en/latest/api/enum/MsoThemeColorIndex.html#msothemecolorindex)枚举的成员来将字体设置为主题颜色：

```python
>>> from docx.enum.dml import MSO_THEME_COLOR
>>> font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
```

可以设置[`ColorFormat`](https://python-docx.readthedocs.io/en/latest/api/dml.html#docx.dml.color.ColorFormat)的 [`rgb`](https://python-docx.readthedocs.io/en/latest/api/dml.html#docx.dml.color.ColorFormat.rgb)或 [`theme_color`](https://python-docx.readthedocs.io/en/latest/api/dml.html#docx.dml.color.ColorFormat.theme_color)属性为**None**将字体的颜色恢复为其默认（继承）值：

```python
>>> font.color.rgb = None
```

确定字体的颜色首先要确定其color type（颜色类型）：

```python
>>> font.color.type
RGB (1)
```

[`type`](https://python-docx.readthedocs.io/en/latest/api/dml.html#docx.dml.color.ColorFormat.type)属性的值可以是[MSO_COLOR_TYPE](https://python-docx.readthedocs.io/en/latest/api/enum/MsoColorType.html#msocolortype)枚举的成员，也可以是**None**。 
*MSO_COLOR_TYPE.RGB*表示它是RGB颜色。 
*MSO_COLOR_TYPE.THEME*指示主题颜色。 
*MSO_COLOR_TYPE.AUTO*指示其值由应用程序自动确定，通常设置为黑色。（此值相对很少。）

**None**表示未应用颜色，并且颜色是从样式层次结构继承的，这是最常见的情况。

当颜色类型为*MSO_COLOR_TYPE.RGB*时， [`rgb`](https://python-docx.readthedocs.io/en/latest/api/dml.html#docx.dml.color.ColorFormat.rgb)属性将是[`RGBColor`](https://python-docx.readthedocs.io/en/latest/api/shared.html#docx.shared.RGBColor)指定的RGB颜色：

```python
>>> font.color.rgb
RGBColor(0x42, 0x24, 0xe9)
```

当颜色类型为*MSO_COLOR_TYPE.THEME*时，[`theme_color`](https://python-docx.readthedocs.io/en/latest/api/dml.html#docx.dml.color.ColorFormat.theme_color)属性将是[MSO_THEME_COLOR_INDEX](https://python-docx.readthedocs.io/en/latest/api/enum/MsoThemeColorIndex.html#msothemecolorindex)的成员所指定主题颜色：

```python
>>> font.color.theme_color
ACCENT_1 (5)
```



------



# Section（节）

Word支持*Section（节）*的概念，即具有相同页面布局设置（例如页边距和页面方向）的文档。例如，文档可以包含纵向布局的某些页面和横向布局的其他页面。

虽然大多数Word文档默认只有单一的*Section*，而且，大多数文档不会去更改默认边距或其他页面布局。但是，当您确实需要更改页面布局时，需要了解*Section*以完成此操作。

## 访问Section

通过 [`Document`](https://python-docx.readthedocs.io/en/latest/api/document.html#docx.document.Document)对象上的sections属性实现对文档sections的访问：

```python
>>> document = Document()
>>> sections = document.sections
>>> sections
<docx.parts.document.Sections object at 0x1deadbeef>
>>> len(sections)
3
>>> section = sections[0]
>>> section
<docx.section.Section object at 0x1deadbeef>
>>> for section in sections:
...     print(section.start_type)
...
NEW_PAGE (2)
EVEN_PAGE (3)
ODD_PAGE (4)
```

从理论上讲，文档中可能没有明确表示出section，至少我还没有看到这种情况。如果您访问的是未知的.docx文件，则可以使用len()检查或尝试IndexError异常。

## 添加一个新Section

**Document.add_section()**方法允许在文档末尾创建新的section。调用此方法后添加的段落和表格将属于这个新的section：

```python
>>> current_section = document.sections[-1]  # last section in document
>>> current_section.start_type
NEW_PAGE (2)
>>> new_section = document.add_section(WD_SECTION.ODD_PAGE)
>>> new_section.start_type
ODD_PAGE (4)
```

## section属性

[`Section`](https://python-docx.readthedocs.io/en/latest/api/section.html#docx.section.Section)对象具有11个属性，这些属性指定页面布局设置。

### section开始类型

[`Section.start_type`](https://python-docx.readthedocs.io/en/latest/api/section.html#docx.section.Section.start_type)描述了该section之前的中断类型：

```python
>>> section.start_type
NEW_PAGE (2)
>>> section.start_type = WD_SECTION.ODD_PAGE
>>> section.start_type
ODD_PAGE (4)
```

start_type的值是[WD_SECTION_START](https://python-docx.readthedocs.io/en/latest/api/enum/WdSectionStart.html#wdsectionstart)枚举的成员。

### 页面尺寸和方向

[`Section`](https://python-docx.readthedocs.io/en/latest/api/section.html#docx.section.Section)中的三个属性描述了页面尺寸和方向。例如，这些可以一起用于将section的方向从纵向更改为横向：

```python
>>> section.orientation, section.page_width, section.page_height
(PORTRAIT (0), 7772400, 10058400)  # (Inches(8.5), Inches(11))
>>> new_width, new_height = section.page_height, section.page_width
>>> section.orientation = WD_ORIENT.LANDSCAPE
>>> section.page_width = new_width
>>> section.page_height = new_height
>>> section.orientation, section.page_width, section.page_height
(LANDSCAPE (1), 10058400, 7772400)
```

### 页边距

[`Section`](https://python-docx.readthedocs.io/en/latest/api/section.html#docx.section.Section)上的七个属性一起指定了各种页面边缘间距，这些间距确定了文本在页面上的显示位置：

```python
>>> from docx.shared import Inches
>>> section.left_margin, section.right_margin
(1143000, 1143000)  # (Inches(1.25), Inches(1.25))
>>> section.top_margin, section.bottom_margin
(914400, 914400)  # (Inches(1), Inches(1))
>>> section.gutter
0
>>> section.header_distance, section.footer_distance
(457200, 457200)  # (Inches(0.5), Inches(0.5))
>>> section.left_margin = Inches(1.5)
>>> section.right_margin = Inches(1)
>>> section.left_margin, section.right_margin
(1371600, 914400)
```



------

# Headers and Footers（页眉和页脚）

Word支持*page headers*（页眉）和*page footers*（页脚）。页眉是出现在每页顶部边缘的文本，与文本主体分开，通常传达上下文信息，例如文档标题，作者，创建日期或页码。文档中的页眉在页面之间是相同的，只有内容上的差别，例如节标题或页码的变化。

页脚在各个方面都类似于页眉，只不过它位于页面底部。请勿将其与footnote（脚注）混淆，脚注在页面之间并不统一。为简便起见，此处使用*header*一词来指代页眉页脚对象，以使读者能够理解其对两种对象类型的适用性。

## 访问Section（节）的header（页眉页脚）

每个section允许有不同的页眉或页脚。例如，横向section的页眉可能比纵向section的页眉宽。

每个section对象都有一个`.header`属性，可用于访问该section的[`_Header`](https://python-docx.readthedocs.io/en/latest/api/section.html#docx.section._Header)对象：

```python
>>> document = Document()
>>> section = document.sections[0]
>>> header = section.header
>>> header
<docx.section._Header object at 0x...>
```

[`_Header`](https://python-docx.readthedocs.io/en/latest/api/section.html#docx.section._Header)对象*始终*存在于Section.header上，即使没有为该section定义页眉也是如此。 
_Header.is_linked_to_previous表示实际的页眉是否被定义：

```python
>>> header.is_linked_to_previous
True
```

值为True表示 [`_Header`](https://python-docx.readthedocs.io/en/latest/api/section.html#docx.section._Header)对象没有页眉定义，并且该section与上个section的页眉相同。这种“继承”行为是递归的，页眉实际上是从具有页眉定义的第一个先前section中获得其定义。在Word中，称为为“与先前相同”。

新文档没有页眉（在包含该标题的单个section上），因此在这种情况下`.is_linked_to_previous`为True。请注意，这种情况可能有点违反直觉，因为没有先前的节头可继承。在这种“没有上一个页眉”的情况下，不显示任何页眉。

## 添加页眉（简单情况）

只需编辑[`_Header`](https://python-docx.readthedocs.io/en/latest/api/section.html#docx.section._Header) 对象的内容，即可将页眉添加到新文档中。 [`_Header`](https://python-docx.readthedocs.io/en/latest/api/section.html#docx.section._Header) 对象的内容就像[`Document`](https://python-docx.readthedocs.io/en/latest/api/document.html#docx.document.Document)对象一样能够被编辑。请注意，就像新文档一样，新标题已经包含一个（空）paragraph（段落）：

```python
>>> paragraph = header.paragraphs[0]
>>> paragraph.text = "Title of my document"
```

![header](https://python-docx.readthedocs.io/en/latest/_images/hdrftr-01.png)

还请注意，添加内容（甚至只是访问header.paragraphs）的行为添加了页眉定义并更改了`.is_linked_to_previous`的状态：

```python
>>> header.is_linked_to_previous
False
```

## 添加“分区”页眉内容

具有多个“区域”的页眉通常是使用精心放置的制表位来完成的。

中心和右对齐“区域”所需的制表位是Word中页眉和页脚样式的一部分。
如果您使用的是自定义模板而不是默认的*python-docx*，则在模板中定义该样式可能很有意义。

插入的制表符（“ \ t”）用于分隔左，中和右对齐的标题内容：

```python
>>> paragraph = header.paragraphs[0]
>>> paragraph.text = "Left Text\tCenter Text\tRight Text"
>>> paragraph.style = document.styles["Header"]
```

![head_zone](https://python-docx.readthedocs.io/en/latest/_images/hdrftr-02.png)

页眉的样式会自动应用到新页眉，因此在这种情况下，代码的第三行（应用页眉样式）是不必要的，但此处包含此行以说明一般情况。

## 移除页眉

可以通过给.is_linked_to_previous属性赋值为True来删除不需要的页眉：

```python
>>> header.is_linked_to_previous = True
>>> header.is_linked_to_previous
True
```

当将True分配给.is_linked_to_previous时，页眉的内容将不可撤消地删除。

## 理解多section文档中的页眉

为了理解多section文档中的页眉行为，一些简单的概念将很有帮助。
这里简而言之：

1. 每个section都可以有自己的页眉定义（但不是必须的）

2. 缺少页眉定义的section将继承前一个section的页眉。_Header.is_linked_to_previous属性仅反映是否存在页眉定义；存在时为False，否则显示为True。
3. 缺少页眉定义是默认状态。新文档还没有定义页眉，新插入的section也没有。.is_linked_to_previous报告在这两种情况下均为True。
4. 如果_Header对象具有定义，则其内容为其自身的内容。如果不是，则其内容为确实具有页眉定义的前一个section页眉的内容。如果没有section具有页眉定义，则在第一个section上添加一个新的页眉定义，其他所有section都继承该页眉。页眉定义发生在第一次访问页眉内容时，通过访问`header.paragraphs`来实现的。

## 添加页眉定义（一般情况）

通过将False分配给.is_linked_to_previous属性，可以为section提供显式的页眉定义：

```python
>>> header.is_linked_to_previous
True
>>> header.is_linked_to_previous = False
>>> header.is_linked_to_previous
False
```

新添加的页眉定义包含一个空的paragraph（段落）。请注意，以这种方式添加页眉定义有时会很有用，因为它可以有效地“关闭”该节的页眉和之后的section，直到下一个具有已定义页眉的section。

在已经具有页眉定义的页眉上将False分配给.is_linked_to_previous不会执行任何操作。

### 继承的内容会自动定位

编辑页眉的内容会更改其“继承”的源页眉的内容。例如，如果第2个section的页眉是从第1个section继承的，并且您编辑了第2个section的页眉，则实际上是在更改第1个section页眉的内容。除非您首先显式地将False分配给它的.is_linked_to_previous属性，否则不会为第2个section添加新的页眉定义。
