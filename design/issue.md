# 问题记录

## 优化
- [x] 可跳过windows隐藏文件
- [x] 筛选不输出文件夹、忽略文件夹、文件后缀黑白名单的功能
- [x] 自动删除del标记的行，可配置
- [x] 文件类型的输出
- [x] 时间戳转化为Excel可用
- [x] 错误的UI提示
- [ ] 可打开文件所在文件夹（通过加一列hyperlink公式实现，os.path.dirname）

## bug
- [x] status记录错误（无误）
- [ ] hyperlink公式未变链接
- [x] 时间戳未存储为Excel时间格式（Excel日期时间有bug，不采用）
- [x] 会产生None|None为key的数据（无法重现，应以解决）
- [x] 文件链接公式需包含起始文件夹
- [x] 文件类型公式需容错未知格式
- [ ] 减小打包文件大小－－使用pipenv
- [x] 三万五+文件运行缓慢，首次1分钟以内，二次内存溢出（加载Excel数据卡死）（改变读取方式，先快速读取数据。已解决，3分钟之内可完成）
- [x] Excel中的日期格式加载错误：datetime格式无法json序列化(使用openpyxl.utils.datetime.to_excel转化存储）
- [ ] Excel中的日期刷新后格式消失
