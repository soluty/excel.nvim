我找了好多neovim中显示excel的插件, 在windows下都不是太好用. 只好自己写一个.

## todo
- autocmd进入excel的时候自动解析.
- 除了第一个table以外还可以查看其它table
- 最好做成二级的, 进入excel文件, 如果有多个表格, 那么给一个表格名字列表出来, 并且列表里的item再点进去就是真正的表格数据
- table的列要显示数字标记, 并且隐藏左边本来的数字, 并且提供去到某行/某列的功能.
- 文本对象需要一个cell对象, 一个line对象和一个column对象.
- 不要暴露全局函数
- 性能情况
