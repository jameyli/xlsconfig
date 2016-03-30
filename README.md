# xlsconfig
游戏配置读取工具

### 主要功能：
1. 配置定义生成，根据excel 自动生成配置的PB定义
2. 配置数据导入，将配置数据生成PB的序列化后的二进制数据或者文本数据

### 说明:
excel 的前四行用于结构定义, 第五行开始为数据
单字段属性定义时:
    * required 必有属性
    * optional 可选属性
    * repeated 重复属性
结构属性定义时:
    * repeated 表示这个结构重复的最大次数
    * struct 表示这个结构的元素个数和结构名

第1行  | required/optional/repeated | repeated  | struct               |
       | ---------------------------| ---------:| --------------------:|
第2行  | 属性类型                   |  最大个数 | 结构元素个数         |
第3行  | 属性名                     |           | 结构类型名           |
第4行  | 注释说明                   |           |                      |
第5行  | 属性值                     |           |                      |


### 内置类型:
* DateTime  "2014-10-14 08:00:00" 自动导出为UNIX时间戳
* TimeDuration "3D05H" 自动导出为 秒数

### 依赖:
* protobuf
* xlrd

### 注意:
* 表名sheet_name 使用大写
* 默认值

## Run Example
```
     cd example
     ../xls_config_tools.py GOODS_CONF goods.xls
```
输出的内容
* .proto 配置的定义
* .data 配置内容 proto 编码后的二进制格式，可以使用proto直接解码
* .txt 配置内容 proto 编码后的明文(json)格式，可以使用proto直接解码
* .py 由 proto 生成的 python 脚本，工具自己生成，可删除
* .log 工具运行日志



