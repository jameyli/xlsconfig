# xlsconfig
游戏配置工具

## Run Example
```
     cd example
     ../xls_config_tools.py GOODS_CONF goods.xls
```
输出的内容
.proto 配置的定义
.data 配置内容 proto 编码后的二进制格式，可以使用proto直接解码
.txt 配置内容 proto 编码后的明文(json)格式，可以使用proto直接解码
.py 由 proto 生成的 python 脚本，工具自己生成，可删除
.log 工具运行日志



