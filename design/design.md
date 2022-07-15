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
  - easygui
- XlManager：实现Excel读写的功能
  - openpyxl
- ConfManager：配置项的加载、读取、存储功能
- FileManager：实际文件信息获取、递归文件目录获取等功能
  - pywin32(win32file,win32con)
- DataManager：数据融合、比较、冲突记录

## 数据结构
- Excel数据

id|名称|数据类型|产生方式|备注
--|--|--|--|--
key|key|string|代码|由其余数据拼接而成，可由用户自定义
status|状态|string|代码|版本状态枚举：new/del/mod
filename|文件名|string|代码|
ext|扩展名|string|代码|
filetype|文件类型|无|代码-公式|在Excel中自定义后映射
is_folder|是文件夹|bool|代码|
path|文件夹路径|string|代码|仅记录相对路径
folder_link|文件夹链接|无|代码-公式|通过path计算
file_link|文件链接|无|代码-公式|通过path计算
size|文件大小|int|代码|
ctime|创建时间|int|代码|Excel浮点数不稳，改用整数
mtime|修改时间|int|代码|
atime|访问时间|int|代码|
c_time|创建时间|string|代码|转化为可读的日期时间格式
m_time|修改时间|int|代码|
a_time|访问时间|int|代码|
label1-n|标签1-n|string|用户维护|自定义项，可使用任意名称，添加任意数量
note|备注|string|用户维护|与label字段作用相同

```python
data = {
  key1:{'label1':'a','label2':'b'...},
  key2:{'label1':'','label2':'a'...},
  ...
}
```

## 配置项

key|名称|类型|默认值|说明
--|--|--|--|--
BASE_FOLDER|起始文件夹|string||递归搜索的文件夹起点，留空即Excel文件所属文件夹(加载后会自动规范化处理)
AUTO_BACKUP|生成列表前自动备份|bool|False|自动备份list文件，生成备份sheet
AUTO_DEL_IGNORED|自动删除被排除文件数据|bool|True|如文件行被标记为del，则自动删除（可能导致部分自定义数据丢失）
NO_HIDDEN_FILES_WIN|排除隐藏文件（Windows）|bool|True|不输出Windows环境下的隐藏文件至列表中
NO_HIDDEN_FILES_POINT|排除.开头的文件及文件夹|bool|True|不输出以.开头的隐藏文件或文件夹至列表中（Linux等环境下的隐藏文件）
NO_FOLDERS|不输出文件夹|bool|False|只输出文件，不输出文件夹至列表中
rIGNORE_KEYWORDS|忽略的关键字|list(Excel的Range)||不输出至列表的绝对路径关键字，存在于文件夹和文件名中的均有效，数量不限
rEXT_BLACKLIST|扩展名黑名单|list(Excel的Range)||文件后缀如在黑名单中，则不会输出至列表中
rEXT_WHITELIST|扩展名白名单|list(Excel的Range)||只有在白名单中的文件后缀，才会输出至列表中
rKEY_MODE|索引构成方式|list(Excel的Range)|[mtime,size]|索引的构成方式，影响确定文件的方法，由key以外的字段拼接而成
rEXT_TO_TYPE|扩展名文件类型映射|list(Excel的Range)||用于vlookup公式映射，代码中不直接引用

## key生成规则
- 遵循用户 KEY_MODE 定义规则，按顺序将定义项的值转化为字符串，并使用`|`拼接。
- 如拼接后的字符串key已存在，则在后面后缀`+n`，从`+1`开始，`+1`存在则`+2`，以此类推。

```py
# 示例
KEY_MODE = ['mtime','size']
mtime = 1651909589
size = 2261

key = '1651909589|2261'
# 如'1651909589|2261'已存在，则
key = '1651909589|2261+1'
# 如'1651909589|2261+1'仍存在，则
key = '1651909589|2261+2'
……
```

## 支持的功能
- list页结构调整：可在key所在行之上任意添加行，可修改key下一行的文字描述，可调换所有字段的前后顺序（需确保key在最前）
- list页可直接加列，添加任意数量的自定义标签用于记录和筛选
- config页结构调整：可任意调整各类config的位置
- 自定义Excel文件及Sheet名称：通过修改`FileManager.json`文件，可自定义Excel文件名、Sheet名
- 每次自动覆盖式备份Excel数据至`.backup.json`，如需还原可转化为csv等使用(https://data.page/json/csv)

## 其余小工具
- filebat：对文件批量移动、重命名、删除等功能的工具。
