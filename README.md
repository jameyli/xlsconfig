# xlsconfig
游戏配置读取工具, 已成功应用于多个项目:v:

### 功能
1. 配置定义生成，根据 excel 自动生成配置的 ProtoBuff(PB) 定义
2. 配置数据导出，将配置数据生成PB的序列化后的二进制数据或者文本(json)数据
3. 只要程序语言支持PB即可

### 优点
* 定义和数据在一起，避免修改不同步
* 基于 PB 自动实现配置数据的读取

### 说明:
#### 每个页签独立定义了一张表
* 表名 sheet_name 必须全部大写, 否则会直接忽略  (适当的限制可以带来更多的自由)

excel 的前四行用于结构定义, 第五行开始为数据
#### 字段属性定义
* required 必有属性
* optional 可选属性
* repeated 重复属性, 即数组, 属性值分号隔开
* struct 结构定义 

|第1行  | required/optional/repeated | struct                        |
|:-----:|:--------------------------:|:-----------------------------:|
|第2行  | 属性类型                   | 结构元素个数*结构重复次数     |
|第3行  | 属性名                     | 属性名                        |
|第4行  | 注释说明                   | 注释                          |
|第5行  | 属性值                     | 空                            |


### 内置类型
* 支持 protobuf 属性类型
* DateTime  配置 "2014-10-14 08:00:00" 自动导出为UNIX时间戳
* TimeDuration 配置 "3D05H" 自动导出为 秒数
* TODO 更多的内置类型

### 依赖:
* protobuf
* xlrd

## Run Example
```
     ../xls_config_tools.py example/goods.xls
```
输出 output 目录, 每个页签独立, 都分别输出以下文件
* .proto 配置的定义, 基于PB，程序可选择静态编译对应的语言，或者运行时动态载入
* .bytes 配置内容 proto 编码后的二进制格式，可以使用对应的PB定义直接解码读取
* .json 配置内容 proto 编码后的明文(json)格式，可以使用对应的PB直接读取
* .py 由 proto 生成的 python 脚本，工具自己生成，可删除
* .log 工具运行日志

## 程序中如何读取？
### C++
```c++
#include <cstdio>
#include <string>
#include "xlsconfig_goods_conf.pb.h"
#include <google/protobuf/text_format.h>

int main(int argc, char* argv[]) {
	const char* data_file = "xlsconfig_goods_conf.data";
	FILE* file = fopen(data_file, "rb");
	assert(file);
	if (!file) {
		return -1;
	}

	static char data[1024 * 1024 * 10];
	size_t readn = fread(data, 1, sizeof(data), file);
	fclose(file);

	xlsconfig::goods::GOODS_CONF_ARRAY conf_array;
	bool parse_ret = conf_array.ParseFromArray(data, readn);
	if (!parse_ret) {
		return -2;
	}

	printf("config load|%s", conf_array.ShortDebugString().c_str());
	return 0;
}
```
### C＃
待补充

### lua
直接输出 lua table

