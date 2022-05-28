中文 | [English](./README_CN.md)

# PyAddin

[![pypi-version](https://img.shields.io/pypi/v/pyaddin.svg)](https://pypi.python.org/pypi/pyaddin/)
![license](https://img.shields.io/pypi/l/pyaddin.svg)

VBA日渐式微，但 Excel 依旧是表格数据存储、处理和传阅的通用工具。对于一些复杂的业务逻辑，一个带宏的 Excel 文件（xlsm）通常包含数据本身和处理数据的VBA脚本。当需要频繁复用这些脚本时，例如按相同的业务逻辑处理月度数据，建议将其拆分为两部分： Excel 纯数据文件（xlsx），以及一个负责数据处理的 Excel 插件（xlam）。

`PyAddin`是一个辅助创建上述 Excel 插件的命令行工具，同时支持使用 Python 语言来实现原本 VBA 负责的数据处理流程。两个基本功能如下：

- 创建一个模板插件，支持自定义菜单功能区（Ribbon）和近“无缝”调用 Python 脚本。
- 在开发插件过程中，方便根据自定义的`CustomUI.xml`更新插件菜单功能区。

集成 VBA 和 Python 的主要思路：VBA 通过后台运行控制台程序调用 Python 脚本，Python 脚本执行计算并以写临时文件的方式保存返回值，最后 VBA 读取返回值。此外，借助 Python 第三方库`pywin32`，可以直接在 Python 脚本中进行与 Excel 的交互，例如获取/设置单元格内容，设置单元格样式等等。


## 限制

- 仅支持 Windows 平台
- 要求 Microsoft Excel 2007 及以上
- VBA 和 Python 之间传参仅限于 **字符串** 格式简单数据类型
- 与 Excel 的交互能力取决于 `pywin32/win32com`


## 安装

支持 `Pypi` 或者本地安装：

```
# pypi
pip install pyaddin

# local  
python setup.py install

# local in development mode
python setup.py develop
```

使用 `pip` 卸载:

```
pip uninstall pyaddin
```

## 命令说明

- 创建模板插件

```
pyaddin init --name=xxx --quiet=True|False
```

- 更新插件功能区

```
pyaddin update --name=xxx --quiet=True|False
```

其中，`quiet`是可选参数，表明是否以后台模式创建插件（不显式打开 Excel）。默认值`True`，即后台模式。

## 使用帮助

### 1. 初始化模板插件

```
D:\WorkSpace>pyaddin init --name=sample
```

在当前目录下新建了文件夹`sample`，其中包含模板插件`sample.xlam`，以及实现VBA 与 Python 互联所需的支持文件。目录结构如下：

```
sample\
|- scripts\
|    |- utils\
|    |    |- __init__.py
|    |    |- context.py
|    |- __init__.py
|    |- sample.py
|- main.cfg
|- main.py
|- CustomUI.xml
|- sample.xlam
```

其中，

- `main.py`是 VBA 调用 Python 脚本的入口文件。
- `main.cfg`是基本配置参数文件，例如指定 Python 解释器的路径。
- `scripts`存放处理具体业务的 Python 脚本，例如`sample.py`是其中的一个示意模块，开发者根据需要在此目录下创建其他模块。
- `CustomUI.xml`定义了插件的 Ribbon 界面，例如包含的控件及样式。
- `sample.xlam`为模板插件，开发者可以在此基础上添加和扩展自定义的功能。

当前模板的 Ribbon 区域参考下图。

![add-in.png](add-in.png)


### 2. 自定义 Ribbon 区域

在上一步创建的 [`CustomUI.xml`](./pyaddin/resources/CustomUI.xml) 的基础上，根据具体需求设计界面和样式，然后运行以下命令将其更新到插件中去。当然， Excel 2007及以上的文件本质上是一个压缩文件，因此可以手动解压后替换相应的`CustomUI.xml`，或者直接借助其他的 Ribbon XML 编辑器进行修改。


```
D:\WorkSpace\sample>pyaddin update --name=sample
```


以此模板为例，定义了两个分组：

- `setting`组提供基础功能，请直接保留在你的项目中。例如，运行前需要设置 Python 解释器路径。
- `Your Group 1`是一个示例分组，其中设计了两个按钮。实际开发中请替换为需要的控件，或者增加其他更多分组。 

关于 Ribbon 界面的具体介绍及格式规范，参考以下链接。

- [General Format of XML Markup Files](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa338202(v%3doffice.12)#general-format-of-xml-markup-files)
- [Custom UI](https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/edc80b05-9169-4ff7-95ee-03af067f35b1)


### 3. 实现 Ribbon 控件响应函数

这一步需要在 VBA 中实现 `CustomUI.xml`定义的控件事件及其响应函数。

以模板插件为例，下面的 XML 片段表明按钮`Sample 1`的点击事件由 VBA 过程`CB_Sample_1`响应。

```xml
<button id="sample1" label="Sample 1" 
		imageMso="AppointmentColor3" size="large" 
		onAction="CB_Sample_1"
...>
```

进一步，查看插件的`UserRibbon`模块中定义的两个响应过程：


```vb
Sub CB_Sample_1(control As IRibbonControl)
    '''onAction for control: Sample 1'''
    Dim res As Object
    Dim x As Integer: x = Range("A1").Value
    Dim y As Integer: y = Range("A2").Value
    
    Set res = RunPython("scripts.sample.run_example_1", x, y)
    Range("A3") = res("value")    
End Sub

Sub CB_Sample_2(control As IRibbonControl)
    '''onAction for control: Sample 2'''
    RunPython "scripts.sample.run_example_2"
End Sub
```

可以看到，二者都是通过`RunPython()`函数来调用 Python 脚本并获取返回值，函数签名如下：

```vb
Function RunPython(methodName As String, ParamArray args()) As Object
```

- 第一个参数`methodName`表示调用的 Python 方法，具体格式为`package.module.method`。本例 "scripts.sample.run_example_1" 表示调用脚本`sample/scripts/sample.py`中的`run_example_1`方法。

- 第二个参数`args`是可变长参数，可以传入任意个数的参数给相应 Python 方法。注意，仅支持字符串、数字等简单数据类型，并且经过控制台参数传递，最终到 Python 端都被转成了字符串格式。

- 返回值为 VBA 的`Dictionary`格式，其中包含两个键：
  - `status`: 调用成功与否，True 或者 False
  - `value`: 返回值（错误信息如果`status`为 False）


### 4. 实现 Python 脚本

根据 VBA 端`RunPython()`调用路径和参数列表，创建相应的 Python 脚本文件及函数。

以模板插件为例，在项目目录下查看`scripts\sample.py`文件：

```python
# sample.py
from .utils import context

def run_example_1(x:str, y:str):
    return int(x) + int(y)

def run_example_2():
    # get the workbook calling this method, then do anything with win32com
    wb = context.get_caller()
    sheet = wb.ActiveSheet

    # get cells value
    x = sheet.Range('A1').Value
    y = sheet.Range('A2').Value

    # set cell value
    sheet.Range('A3').Value = x + y
```

根据上下文可知，上述两个方法实现相同的事情，将单元格 A1、A2 的和写到单元格 A3，进一步可以分为三步：取值，求和，设置值。但实现的思路略有不同：

- `run_example_1`的思路是仅将 Excel 无关的流程（求和）交由 Python 实现，所有 Excel 相关的操作（取值、设置值）仍通过 VBA 执行，二者的桥梁是参数传递。

- `run_example_2`的思路是在 Python 端获取到调用此脚本的工作簿，然后直接对其做所有需要的操作（取值，求和，设置值）。

对比可以发现，前者按各自强项分工，运行效率较高，但 VBA 和 Python 之间存在强耦合，且容易受限于参数传递的复杂度；后者几乎完全解耦 VBA 和 Python，且避免了参数传递，但操作 Excel 的能力受限于`pywin32`，某些操作可能无法实现。


### 5. 交付插件

因为依赖关系，整个工程应作为整体（第一步中所示结构）交付，且需要最终用户配置好 Python 环境及相应第三方库（主要是`pywin32`）。有时候 Python 环境对用户来说是一个挑战，因此可以考虑可移植的便携式 Python，事先安装好依赖库后随插件工程一起发布，如此便于最终用户开箱即用。


## 许可

MIT License