---
export_on_save:
    html: true
---

## 流程

```mermaid
graph TB

点击Python程序-->检查Excel文件{检查Excel文件}
检查Excel文件--有效-->加载配置
检查Excel文件--无效-->生成Excel文件-->加载配置
加载配置-->加载Excel数据
加载Excel数据-.->Excel数据[(Excel数据)]
加载Excel数据-->解析递归文件信息
解析递归文件信息-.->文件数据[(文件数据)]
解析递归文件信息-->数据融合
Excel数据[(Excel数据)]-.->数据融合
文件数据[(文件数据)]-.->数据融合
数据融合-->冲突记录
冲突记录-->写入Excel数据
```

## 脚本结构
- Main：主流程
- XlManager：实现Excel读写的功能
  - openpyxl
- ConfManager：配置项的加载、读取、存储功能
  - os
  - json
  - codes
- FileManager：实际文件信息获取、递归文件目录获取等功能
  - os
- DataManager：数据融合、比较、冲突记录

## 数据结构
- Excel数据

id|名称|数据类型|产生方式|备注
--|--|--|--|--
key|key|string|代码|由其余数据拼接而成，可由用户自定义
label1|标签1|string|用户维护|自定义项
label2|标签2|string|用户维护|自定义项
label3|标签3|string|用户维护|自定义项
filename|文件名|string|代码|
ext|扩展名|string|代码|
filetype|文件类型|无|公式|在Excel中自定义后映射
is_folder|是文件夹|bool|代码|
path|文件路径|string|代码|仅记录相对路径
hyperlink|文件链接|无|公式|通过path计算
size|文件大小|int|代码|
ctime|创建时间|float|代码|
mtime|修改时间|float|代码|
atime|访问时间|float|代码|
note|备注|string|用户维护|

```python
data = {
  key1:{'label1':'a','label2':'b'...},
  key2:{'label1':'','label2':'a'...},
  ...
}
```

## 配置项

key|名称|类型|默认值|说明
--|--|--|--|--|--
BASE_FOLDER|起始文件夹|string||递归搜索的文件夹起点，留空即Excel文件所属文件夹
AUTO_BACKUP|生成列表前自动备份|bool|True|自动备份list文件，生成备份sheet
NO_HIDDEN_FILES_WIN|排除隐藏文件（Windows）|bool|True|不输出Windows环境下的隐藏文件至列表中
NO_HIDDEN_FILES_POINT|排除.开头的文件及文件夹|bool|True|不输出以.开头的隐藏文件或文件夹至列表中（Linux等环境下的隐藏文件）
NO_FOLDERS|不输出文件夹|bool|False|只输出文件，不输出文件夹至列表中
rIGNORE_FOLDERS|忽略的文件夹|list(Excel的Range)||不输出至列表的文件夹，数量不限。如绝对路径中存在其中内容，则会跳过
rEXT_BLACKLIST|文件后缀黑名单|list(Excel的Range)||文件后缀如在黑名单中，则不会输出至列表中
rEXT_WHITELIST|文件后缀白名单|list(Excel的Range)||只有在白名单中的文件后缀，才会输出至列表中
rKEY_MODE|索引构成方式|list(Excel的Range)|[mtime,size]|索引的构成方式，影响确定文件的方法，由key以外的字段拼接而成

## key生成规则
- 遵循用户 KEY_MODE 定义规则，按顺序将定义项的值转化为字符串，并使用`|`拼接。
- 如拼接后的字符串key已存在，则在后面后缀`+n`，从`+1`开始，`+1`存在则`+2`，以此类推。

```py
# 示例
KEY_MODE = ['mtime','size']
mtime = 1651909589.183366
size = 2261

key = '1651909589.183366|2261'
# 如'1651909589.183366|2261'已存在，则
key = '1651909589.183366|2261+1'
# 如'1651909589.183366|2261+1'仍存在，则
key = '1651909589.183366|2261+2'
……
```

## 支持的功能
- list页结构调整：可在key所在行之上任意添加行，可修改key下一行的文字描述，可调换所有字段的前后顺序（需确保key在最前）
- config页结构调整：可任意调整各类config的位置
- 自定义Excel文件及Sheet名称：通过修改json文件
