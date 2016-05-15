#! /usr/bin/env python
#coding=utf-8

##
# @file:   xls_config_tool.py
# @author: jameyli <github.com/jameyli>
# @brief:  xls 配置导表工具

#  import xlrd # for read excel
import sys
import os
import subprocess
import getopt
import shutil
import re
import pandas as pd
import datetime
import time

reload(sys)
sys.setdefaultencoding( "utf-8" )

PROTOC_BIN = "protoc"
LUA_BIN = "lua"

OUTPUT_PATH = "output/"
PROTO_OUTPUT_PATH = OUTPUT_PATH + "proto/"
BYTES_OUTPUT_PATH = OUTPUT_PATH + "bytes/"
TEXT_OUTPUT_PATH = OUTPUT_PATH + "text/"
LUA_OUTPUT_PATH = OUTPUT_PATH + "lua/"

CPP_OUTPUT_PATH = ""#PROTO_OUTPUT_PATH + "cpp/"
PYTHON_OUTPUT_PATH = ""#PROTO_OUTPUT_PATH + "python/"

OUTPUT_FILE_PREFIX = "xlsc_"

INTEGER_TYPES = ["int32", "int64", "uint32", "uint64"]
FRACTION_TYPES = ["float", "double"]

# TAP的空格数
TAP_BLANK_NUM = 4

FIELD_RULE_ROW = 0
# 这一行还表示重复的最大个数，或结构体元素数
FIELD_TYPE_ROW = 1
FIELD_NAME_ROW = 2
FIELD_COMMENT_ROW = 3
DATA_BEGIN_ROW = 4

###############################################################################
# 日志相关
class Color :
    BLACK = "\033[30m"
    RED = "\033[31m"
    GREEN = "\033[32m"
    YELLOW = "\033[33m"
    BLUE = "\033[34m"
    PURPLE = "\033[35m"
    CYAN = "\033[36m"
    WHITE = "\033[37m"
    NONE = "\033[0m"



###############################################################################
def GetValue(typename, value_str) :
    #  print typename, value_str, type(value_str)
    if len(str(value_str).strip()) <=0 :
        return None

    if value_str == "nan" :
        return None

    if typename in INTEGER_TYPES :
        return int(value_str)
    elif typename in FRACTION_TYPES :
        return float(value_str)
    elif typename == "DateTime" :
        typename = unicode(value_str).encode('utf-8')
        time_struct = time.strptime(value_str, "%Y-%m-%d %H:%M:%S")
        #  #GMT time
        timt_stamp = int(time.mktime(time_struct) - time.timezone)
        return timt_stamp
    elif typename == "TimeDuration" :
        field_value = unicode(value_str).encode('utf-8')
        time_struct=0
        try :
            time_struct = time.strptime(value_str, "%HH")
        except BaseException, e :
            time_struct = time.strptime(value_str, "%jD%HH")
        return 3600 * (time_struct.tm_yday * 24 + time_struct.tm_hour)
    elif typename == "bool" :
        return (True if int(value_str) != 0 else False)
    else :
        return unicode(value_str)

def GetLuaValue(typename, value_str) :
    value = GetValue(typename, value_str)
    if value == None :
        return None

    if typename == "string" :
        return "\"" + str(value) + "\""
    return value

def GetDefaultValue(typename) :
    if typename in INTEGER_TYPES or typename in FRACTION_TYPES or typename == "DateTime" or typename == "TimeDuration" :
        return 0
    elif typename == "bool" :
        return False
    elif typename == "string" :
        return '\"\"'
    else :
        return None

###############################################################################

class StructItem:
    def __init__(self):
        self.field_num = 1
        self.repeated_count = 1
        self.struct_name = ""

class FieldItem:
    def __init__(self):
        self.rule = ""
        self.typename = ""
        self.layout_typename = ""
        self.name = ""
        self.comment = ""
        self.default_value_str = ""
        self.default_value = None

        self.group = None
        self.struct = None

        self.value_str = ""
        self.value = None
    def GetValue(self) :
        return GetValue(self.typename, self.value_str)

def GetField(sheet, col) :
    #创建对象。方便后面访问
    field = FieldItem()
    field.rule = str(sheet.get_value(FIELD_RULE_ROW, col)).strip()
    if ('=' in field.rule) :
        tmp_list = field.rule.split('=')
        field.rule = tmp_list[0].strip()
        field.group = tmp_list[1].strip()

    field.typename = str(sheet.get_value(FIELD_TYPE_ROW, col)).strip()
    field.layout_typename = field.typename
    if field.typename in ["DateTime", "TimeDuration"] :
        field.layout_typename = "uint32"
    field.name = str(sheet.get_value(FIELD_NAME_ROW, col)).strip().strip()
    if ('=' in field.name) :
        tmp_list = field.name.split('=')
        field.name = tmp_list[0].strip()
        field.default_value_str = tmp_list[1].strip()
        field.default_value = GetValue(field.typename, field.default_value_str)

    field.comment = unicode(sheet.get_value(FIELD_COMMENT_ROW, col)).encode("utf-8")

    if field.rule == "struct" :
        field.struct = StructItem()
        field.struct.field_num = int(str(sheet.get_value(FIELD_TYPE_ROW, col)).split('*')[0])
        field.struct.repeated_num = int(str(sheet.get_value(FIELD_TYPE_ROW, col)).split('*')[1])
        field.struct.struct_name = "InternalType_" + field.name;
        field.typename = field.struct.struct_name
        field.layout_typename = field.typename

    return field


###############################################################################

class SheetInterpreter:
    """通过excel配置生成配置的protobuf定义文件"""
    def __init__(self, xls_file_path, sheet_name, sheet, group):
        self._xls_file_path = xls_file_path
        self._sheet = sheet
        self._sheet_name = sheet_name

        self._group = group

        self._row_count = len(self._sheet.index)
        self._col_count = len(self._sheet.columns)

        #  print self._row_count, self._col_count

        self._row = 0
        self._col = 0

        # 将所有的输出先写到一个list， 最后统一写到文件
        self._output = []
        self._key = []
        # 排版缩进空格数
        self._indentation = 0
        # field number 结构嵌套时使用列表
        # 新增一个结构，行增一个元素，结构定义完成后弹出
        self._field_index_list = [1]
        # 当前行是否输出，避免相同结构重复定义
        self._is_layout = True
        # 保存所有结构的名字
        self._struct_name_list = []

        self._package_name = os.path.splitext(os.path.basename(self._xls_file_path))[0].lower()
        self._pb_file_name = PROTO_OUTPUT_PATH + OUTPUT_FILE_PREFIX + self._package_name + "_" + self._sheet_name.lower() + ".proto"


    def Interpreter(self) :
        """对外的接口"""

        self._LayoutFileHeader()

        package_name_line = "package xlsc." + self._package_name + ";\n"
        self._output.append(package_name_line)

        self._LayoutStructHead(self._sheet_name)
        self._IncreaseIndentation()

        while self._col < self._col_count :
            self._FieldDefine()

        self._DecreaseIndentation()
        self._LayoutStructTail()

        self._LayoutArray()

        self._Write2File()

        # 将PB转换成py格式
        try :
            subprocess.call([PROTOC_BIN, "--python_out=.", self._pb_file_name])
        except BaseException, e :
            print "protoc failed!"
            raise

    def _FieldDefine(self) :
        field = GetField(self._sheet, self._col)
        #  print field.rule

        if field.rule == "key" :
            field.rule = "required"
            #  self._key.append(" "*self._indentation + field.rule + " " + field.typename \
            #          + " " + field.name + " = " + self._GetAndAddFieldIndex() + ";\n")

        if field.rule in ["required", "optional", "repeated"] :
            self._col += 1
            if self._group != None and field.group != None and (field.group not in self._group) :
                return
            self._LayoutComment(field.comment)
            self._LayoutOneField(field)

        elif field.rule == "struct":
            if (self._IsStructDefined(field.struct.struct_name)) :
                self._is_layout = False
            else :
                self._struct_name_list.append(field.struct.struct_name)
                self._is_layout = True

            self._col += 1
            #  if self._group != None and field.group != None and (field.group not in self._group) :
            #      self._is_layout = False

            col_begin = self._col
            self._StructDefine(field.struct.struct_name, field.struct.field_num, field.comment)
            col_end = self._col

            self._col += (field.struct.repeated_num-1) * (col_end-col_begin)

            self._is_layout = True
            #  if self._group != None and field.group != None and (field.group not in self._group) :
            #      return

            field.rule = "repeated"
            self._LayoutOneField(field)

        else :
            self._col += 1
            return

    def _IsStructDefined(self, struct_name) :
        return struct_name in self._struct_name_list

    def _StructDefine(self, struct_name, field_num, comment) :
        """嵌套结构定义"""

        self._LayoutComment(comment)
        self._LayoutStructHead(struct_name)
        self._IncreaseIndentation()
        self._field_index_list.append(1)

        for i in range(field_num) :
            self._FieldDefine()

        self._field_index_list.pop()
        self._DecreaseIndentation()
        self._LayoutStructTail()

    def _LayoutFileHeader(self) :
        """生成PB文件的描述信息"""
        self._output.append("/**\n")
        self._output.append("* @file: " + self._pb_file_name + "\n")
        self._output.append("* @note: Generated by xlsconfig (see https://github.com/jameyli/xlsconfig).  DO NOT EDIT!\n")
        self._output.append("* @source: xls ("+ self._xls_file_path+") sheet ("+ self._sheet_name + ")\n")
        self._output.append("**/\n\n")

    def _LayoutStructHead(self, struct_name) :
        """生成结构头"""
        if not self._is_layout :
            return
        self._output.append("\n")
        self._output.append(" "*self._indentation + "message " + struct_name + "{\n")

    def _LayoutStructTail(self) :
        """生成结构尾"""
        if not self._is_layout :
            return
        self._output.append(" "*self._indentation + "}\n")
        self._output.append("\n")

    def _LayoutComment(self, comment) :
        # 改用C风格的注释，防止会有分行
        if not self._is_layout :
            return
        if comment.count("\n") > 1 :
            if comment[-1] != '\n':
                comment = comment + "\n"
                comment = comment.replace("\n", "\n" + " " * (self._indentation + TAP_BLANK_NUM),
                        comment.count("\n")-1 )
                self._output.append(" "*self._indentation + "/** " + comment + " "*self._indentation + "*/\n")
        else :
            self._output.append(" "*self._indentation + "/** " + comment + " */\n")

    def _LayoutOneField(self, field) :
        """输出一行定义"""
        if not self._is_layout :
            return

        if field.default_value == None:
            self._output.append(" "*self._indentation + field.rule + " " + field.layout_typename \
                    + " " + str(field.name) + " = " + self._GetAndAddFieldIndex() + ";\n")
        else :
            self._output.append(" "*self._indentation + field.rule + " " + field.layout_typename \
                    + " " + str(field.name) + " = " + self._GetAndAddFieldIndex()\
                    + " [default = " + str(field.default_value) + "]" + ";\n")
        return

    def _IncreaseIndentation(self) :
        """增加缩进"""
        self._indentation += TAP_BLANK_NUM

    def _DecreaseIndentation(self) :
        """减少缩进"""
        self._indentation -= TAP_BLANK_NUM

    def _GetAndAddFieldIndex(self) :
        """获得字段的序号, 并将序号增加"""
        index = str(self._field_index_list[- 1])
        self._field_index_list[-1] += 1
        return index

    def _LayoutArray(self) :
        """输出数组定义"""
        self._output.append("message " + self._sheet_name + "_ARRAY {\n")
        self._output.append("    repeated " + self._sheet_name + " items = 1;\n}\n")

    def _Write2File(self) :
        """输出到文件"""
        if not os.path.exists(PROTO_OUTPUT_PATH) : os.makedirs(PROTO_OUTPUT_PATH)
        pb_file = open(self._pb_file_name, "w+")
        pb_file.writelines(self._output)
        pb_file.close()

###############################################################################

class DataParser:
    """解析excel的数据"""
    def __init__(self, xls_file_path, sheet_name, sheet, group):
        self._xls_file_path = xls_file_path
        self._sheet = sheet
        self._sheet_name = sheet_name

        self._group = group

        self._row_count = len(self._sheet.index)
        self._col_count = len(self._sheet.columns)

        self._row = 0
        self._col = 0

        self._package_name = os.path.splitext(os.path.basename(self._xls_file_path))[0].lower()
        self._data_file_name = BYTES_OUTPUT_PATH + OUTPUT_FILE_PREFIX + self._package_name + "_" + self._sheet_name.lower() + ".bytes"
        self._module_name = OUTPUT_FILE_PREFIX + self._package_name + "_" + self._sheet_name.lower() + "_pb2"

        try:
            sys.path.append(PROTO_OUTPUT_PATH)
            exec('from '+self._module_name + ' import *');
            self._module = sys.modules[self._module_name]
        except BaseException, e :
            print "load module(%s) failed"%(self._module_name)
            raise

    def Parse(self) :
        """对外的接口:解析数据"""

        item_array = getattr(self._module, self._sheet_name+'_ARRAY')()

        # 先找到定义ID的列
        id_col = 0
        for id_col in range(self._col_count) :
            info_id = str(self._sheet.get_value(self._row, id_col)).strip()
            if info_id == "" :
                continue
            else :
                break

        for self._row in range(DATA_BEGIN_ROW, self._row_count) :
            # 如果 id 是 空 直接跳过改行
            info_id = str(self._sheet.get_value(self._row, id_col)).strip()
            if info_id == "" :
                continue
            item = item_array.items.add()
            self._ParseLine(item)

        self._WriteReadableData2File(str(item_array))

        #print item_array

        data = item_array.SerializeToString()
        self._WriteData2File(data)

    def _ParseLine(self, item) :

        self._col = 0
        while self._col < self._col_count :
            self._ParseField(item)

    def _ParseField(self, item) :
        field = GetField(self._sheet, self._col)
        field.value_str = str(self._sheet.get_value(self._row, self._col))

        if field.rule == "key" :
            field.rule = "required"

        if field.rule in ["required", "optional"]:
            self._col += 1
            if self._group != None and field.group != None and (field.group not in self._group) :
                return

            field.value = field.GetValue()
            if field.value != None :
                item.__setattr__(field.name, field.value)

        elif field.rule == "repeated" :
            self._col += 1
            if self._group != None and field.group != None and (field.group not in self._group) :
                return

            field_value_list = []
            if len(field.value_str) > 0:
                field_value_list = field.value_str.strip().split(";")
                for value in field_value_list :
                    item.__getattribute__(field.name).append(GetValue(field.typename, value))

        elif "struct" in field.rule :
            self._col += 1

            # 至少循环一次
            if field.struct.repeated_num < 1 :
                field.struct.repeated_num = 1

            for i in range(field.struct.repeated_num):
                struct_item = item.__getattribute__(field.name).add()
                self._ParseStruct(field.struct.field_num, struct_item)
                if self._group != None and field.group != None and (field.group not in self._group) :
                    item.__getattribute__(field.name).__delitem__(-1)
                    continue
                if len(struct_item.ListFields()) == 0 :
                    item.__getattribute__(field.name).__delitem__(-1)
        else :
            self._col += 1
            return

    def _ParseStruct(self, field_num, struct_item) :
        """嵌套结构数据读取"""

        # 跳过结构体定义
        for i in range(field_num) :
            self._ParseField(struct_item)

    def _WriteData2File(self, data) :
        if not os.path.exists(BYTES_OUTPUT_PATH) : os.makedirs(BYTES_OUTPUT_PATH)
        file = open(self._data_file_name, 'wb+')
        file.write(data)
        file.close()

    def _WriteReadableData2File(self, data) :
        if not os.path.exists(TEXT_OUTPUT_PATH) : os.makedirs(TEXT_OUTPUT_PATH)
        file_name = TEXT_OUTPUT_PATH + OUTPUT_FILE_PREFIX + self._package_name + "_" + self._sheet_name.lower() + ".text"
        file = open(file_name, 'wb+')
        file.write(data)
        file.close()

###############################################################################

class LuaParser:
    """excel to lua"""
    def __init__(self, xls_file_path, sheet_name, sheet, group):
        self._xls_file_path = xls_file_path
        self._sheet = sheet
        self._sheet_name = sheet_name

        self._row_count = len(self._sheet.index)
        self._col_count = len(self._sheet.columns)

        self._group = group

        self._row = 0
        self._col = 0

        self.row_key = {}

        self.all_str = ""
        self.row_str = ""


        self._package_name = os.path.splitext(os.path.basename(self._xls_file_path))[0].lower()
        self._data_file_name = LUA_OUTPUT_PATH + OUTPUT_FILE_PREFIX + self._package_name + "_" + self._sheet_name.lower() + ".lua"

    def Parse(self) :
        self.all_str = "-- Generated by xlsconfig (see https://github.com/jameyli/xlsconfig).  DO NOT EDIT!\n"
        self.all_str += "-- source: xls ("+ self._xls_file_path+") sheet ("+ self._sheet_name + ")\n\n"
        self.all_str += self._sheet_name + "={\n"
        for self._row in range(DATA_BEGIN_ROW, self._row_count) :
            self._ParseLine()
            self.all_str += self.row_str + ",\n"
        self.all_str += "}"

        result = re.sub(',}', '}', self.all_str)

        self._WriteData2File(result)
        self._CheckLua()

    def _ParseLine(self) :
        self._col = 0
        self.row_key = {}
        self.row_str = "{"
        while self._col < self._col_count :
            self._ParseField()

        self.row_str += "}"

        key_str = ""
        if (len(self.row_key)) > 1 :
            key_str = "\""
            for key in self.row_key.values() :
                key_str += str(key)
            key_str += "\""
        elif (len(self.row_key)) == 1 :
            key_str += str(self.row_key.values()[0])

        if len(key_str) > 0 :
            self.row_str = "[" + key_str + "] = " + self.row_str

        #  print self.row_str

    def _ParseField(self) :
        dict_item = {}
        field = GetField(self._sheet, self._col)
        field.value_str = str(self._sheet.get_value(self._row, self._col))

        if field.rule == "key" :
            field.value = GetLuaValue(field.typename, field.value_str)
            self.row_key[field.name] = field.value

        if field.rule in ["key", "required", "optional"] :
            self._col += 1
            if self._group != None and field.group != None and (field.group not in self._group) :
                return

            field.value = GetLuaValue(field.typename, field.value_str)
            if field.value != None :
                self.row_str += field.name + '=' + str(field.value) + ','
            elif field.default_value != None :
                self.row_str += field.name + '=' + str(field.default_value) + ','


        elif field.rule == "repeated" :
            self._col += 1
            if self._group != None and field.group != None and (field.group not in self._group) :
                return

            field_value_list = []
            if len(field.value_str) > 0:
                field_value_list = field.value_str.strip().split(";")

                vstr = field.name + '= {'
                for e in field_value_list :
                    v = GetLuaValue(field.typename, e)
                    vstr += '' + str(v) + ','
                self.row_str += vstr + '},'


        elif field.rule == "struct":
            self._col += 1
            if self._group != None and field.group != None and (field.group not in self._group) :
                self._col += field.struct.field_num * field.struct.repeated_num
                return

            self.row_str += field.name + "={"
            # 至少循环一次
            if field.struct.repeated_num < 1 :
                field.struct.repeated_num = 1

            for i in range(field.struct.repeated_num):
                self._ParseStruct(field.struct.field_num)
            self.row_str += "},"
        else :
            self._col += 1
            return

    def _ParseStruct(self, field_num) :
        """嵌套结构数据读取"""

        # 跳过结构体定义
        self.row_str += "{"
        for i in range(field_num) :
            self._ParseField()
        self.row_str += "},"

    def _WriteData2File(self, data) :
        if not os.path.exists(LUA_OUTPUT_PATH) : os.makedirs(LUA_OUTPUT_PATH)
        file = open(self._data_file_name, 'wb+')
        file.write(data)
        file.close()

    def _CheckLua(self) :
        status = subprocess.call([LUA_BIN, self._data_file_name])
        if (status != 0) :
            print Color.RED + "[ERROR]: Test " + self._data_file_name + "  FAILED!" + Color.NONE
            raise

###############################################################################

def ProcessOneFile(xls_file_path, output, group) :
    if not ".xls" in xls_file_path :
        return

    try :
        workbook = pd.ExcelFile(xls_file_path)
    except BaseException, e :
        print Color.RED + "Open %s failed! File is NOT exist!" %(xls_file_path) + Color.NONE
        sys.exit(-2)

    sheet_name_list = workbook.sheet_names

    for sheet_name in sheet_name_list :
        if (not sheet_name.isupper()):
            print "Skip %s" %(sheet_name)
            continue

        if ("=" in sheet_name) :
            sheet_name = sheet_name.split('=')[0]
            sheet_group = sheet_name.split('=')[1]
            if group != None and sheet_group != None and (sheet_group not in group) :
                print "Skip %s" %(sheet_name)
                continue

        sheet = workbook.parse(sheet_name, header=None)
        if ("T." in sheet_name) :
            sheet = sheet.T
            sheet_name = sheet_name.split('.')[1]

        if output == None :
            output = ["proto", "bytes", "lua"]

        if "proto" in output or "bytes" in output:
            interpreter = SheetInterpreter(xls_file_path, sheet_name, sheet, group)
            interpreter.Interpreter()
            print Color.GREEN + "Parse %s of %s TO %s Success!!!" %(sheet_name, xls_file_path, interpreter._pb_file_name) + Color.NONE

        if "bytes" in output:
            parser = DataParser(xls_file_path, sheet_name, sheet, group)
            parser.Parse()
            print Color.GREEN + "Parse %s of %s TO %s Success!!!" %(sheet_name, xls_file_path, parser._data_file_name) + Color.NONE

        if "lua" in output:
            parser = LuaParser(xls_file_path, sheet_name, sheet, group)
            parser.Parse()
            print Color.GREEN + "Parse %s of %s TO %s Success!!!" %(sheet_name, xls_file_path, parser._data_file_name) + Color.NONE

    workbook.close()

def ProcessPath(file_path, output, group) :
    if os.path.isfile(file_path) :
        ProcessOneFile(file_path, output, group)
    elif os.path.isdir(file_path) :
        path_list = os.listdir(file_path)
        for path in path_list :
            real_file_path = file_path+"/"+path
            ProcessPath(real_file_path, output, group)

def usage():
    print '''
Usage: %s [options] excel_file
option:
    -h, --help
    -o, --output=   proto, bytes, lua
    -g, --group=    S, C or user defined in excel like "optional = S"
''' %(sys.argv[0])


if __name__ == '__main__' :
    try:
        opt, args = getopt.getopt(sys.argv[1:], "ho:g:", ["help", "output=", "group="])
    except getopt.GetoptError, err:
        print "err:",(err)
        usage()
        sys.exit(-1)

    if len(args) > 0 :
        xls_file_path = args[0]
    else :
        xls_file_path = "."

    output = None
    group = None
    for op, value in opt:
        if op == "-h" or op == "--help":
            usage()
            sys.exit(0)
        elif op == "-o" or op == "--output":
            output = []
            output.append(value)
        elif op == "-g" or op == "--group":
            group = []
            group.append(value)

    if os.path.exists(OUTPUT_PATH) : shutil.rmtree(OUTPUT_PATH)


    ProcessPath(xls_file_path, output, group)


