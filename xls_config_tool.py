#! /usr/bin/env python
#coding=utf-8

##
# @file:   xls_config_tool.py
# @author: jameyli <lgy AT live DOT com>
# @brief:  xls 配置导表工具

import xlrd # for read excel
import sys
import os

reload(sys)
sys.setdefaultencoding( "utf-8" )

# TAP的空格数
TAP_BLANK_NUM = 4

FIELD_RULE_ROW = 0
# 这一行还表示重复的最大个数，或结构体元素数
FIELD_TYPE_ROW = 1
FIELD_NAME_ROW = 2
FIELD_COMMENT_ROW = 3
DATA_BEGIN_ROW = 4

PROTOC_BIN = "protoc "

OUTPUT_PATH = "" #output/"
PROTO_OUTPUT_PATH = ""#OUTPUT_PATH + "proto/"
BYTES_OUTPUT_PATH = ""#OUTPUT_PATH + "bytes/"
JSON_OUTPUT_PATH = ""#OUTPUT_PATH + "json/"
LUA_OUTPUT_PATH = ""#OUTPUT_PATH + "lua/"

CPP_OUTPUT_PATH = ""#PROTO_OUTPUT_PATH + "cpp/"
PYTHON_OUTPUT_PATH = ""#PROTO_OUTPUT_PATH + "python/"

OUTPUT_FULE_PATH_BASE = "xlsc_"

DIGITAL_TYPES = ["int32", "int64", "uint32", "uint64", "double", "float"]

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

class LogHelp :
    """日志辅助类"""
    _logger = None
    _close_imme = True

    @staticmethod
    def set_close_flag(flag):
        LogHelp._close_imme = flag

    @staticmethod
    def _initlog():
        import logging

        LogHelp._logger = logging.getLogger()
        logfile = 'xls_config_tool.log'
        hdlr = logging.FileHandler(logfile)
        formatter = logging.Formatter('%(asctime)s|%(levelname)s|%(lineno)d|%(funcName)s|%(message)s')
        hdlr.setFormatter(formatter)
        LogHelp._logger.addHandler(hdlr)
        LogHelp._logger.setLevel(logging.NOTSET)
        # LogHelp._logger.setLevel(logging.WARNING)

        LogHelp._logger.info("\n\n\n")
        LogHelp._logger.info("logger is inited!")

    @staticmethod
    def get_logger() :
        if LogHelp._logger is None :
            LogHelp._initlog()

        return LogHelp._logger

    @staticmethod
    def close() :
        if LogHelp._close_imme:
            import logging
            if LogHelp._logger is None :
                return
            logging.shutdown()

# log macro
LOG_DEBUG=LogHelp.get_logger().debug
LOG_INFO=LogHelp.get_logger().info
LOG_WARN=LogHelp.get_logger().warn
LOG_ERROR=LogHelp.get_logger().error
###############################################################################

class StuctItem:
    def __init__(self):
        self.field_num = 1
        self.repeated_count = 1
        self.struct_name = ""

class FieldItem:
    def __init__(self):
        self.rule = ""
        self.typename = ""
        self.name = ""
        self.comment = ""
        self.default_value = ""
        self.group = None
        self.struct = None

class SheetInterpreter:
    """通过excel配置生成配置的protobuf定义文件"""
    def __init__(self, xls_file_path, sheet):
        self._xls_file_path = xls_file_path
        self._sheet = sheet
        self._sheet_name = self._sheet.name

        # 行数和列数
        self._row_count = len(self._sheet.col_values(0))
        self._col_count = len(self._sheet.row_values(0))

        self._row = 0
        self._col = 0

        # 将所有的输出先写到一个list， 最后统一写到文件
        self._output = []
        # 排版缩进空格数
        self._indentation = 0
        # field number 结构嵌套时使用列表
        # 新增一个结构，行增一个元素，结构定义完成后弹出
        self._field_index_list = [1]
        # 当前行是否输出，避免相同结构重复定义
        self._is_layout = True
        # 保存所有结构的名字
        self._struct_name_list = []

        self._pb_file_name = OUTPUT_FULE_PATH_BASE + self._sheet_name.lower() + ".proto"


    def Interpreter(self) :
        """对外的接口"""
        LOG_INFO("begin Interpreter, row_count = %d, col_count = %d", self._row_count, self._col_count)

        self._LayoutFileHeader()

        package_name_line = "package xlsc." + os.path.splitext(os.path.basename(self._xls_file_path))[0] + ";\n"
        self._output.append(package_name_line)

        self._LayoutStructHead(self._sheet_name)
        self._IncreaseIndentation()

        while self._col < self._col_count :
            self._FieldDefine()

        self._DecreaseIndentation()
        self._LayoutStructTail()

        self._LayoutArray()

        self._Write2File()

        LogHelp.close()

        # 将PB转换成py格式
        try :
            command = PROTOC_BIN + " --python_out=./ " + PROTO_OUTPUT_PATH + self._pb_file_name
            os.system(command)
        except BaseException, e :
            print "protoc failed!"
            raise

    def _FieldDefine(self) :
        field = FieldItem()
        field.rule = str(self._sheet.cell_value(FIELD_RULE_ROW, self._col))
        field.typename = str(self._sheet.cell_value(FIELD_TYPE_ROW, self._col)).strip()
        field.name = str(self._sheet.cell_value(FIELD_NAME_ROW, self._col)).strip()
        field.comment = unicode(self._sheet.cell_value(FIELD_COMMENT_ROW, self._col)).encode("utf-8")

        LOG_INFO("row=%d, col=%d|%s|%s|%s|%s", self._row, self._col,field.rule, field.typename, field.name, field.comment)

        if field.rule in ["required", "optional", "repeated"] :
            self._LayoutComment(field.comment)
            self._LayoutOneField(field.rule, field.typename, field.name)
            self._col += 1

        elif field.rule == "struct":
            field_num = int(str(self._sheet.cell_value(FIELD_TYPE_ROW, self._col)).split('*')[0])
            repeated_num = int(str(self._sheet.cell_value(FIELD_TYPE_ROW, self._col)).split('*')[1])
            struct_name = "InternalType_" + field.name;

            LOG_INFO("%s|%d|%s|%s", field.rule, field_num, struct_name, field.name)

            if (self._IsStructDefined(struct_name)) :
                self._is_layout = False
            else :
                self._struct_name_list.append(struct_name)
                self._is_layout = True

            self._col += 1

            col_begin = self._col
            self._StructDefine(struct_name, field_num, field.comment)
            col_end = self._col

            field.rule = "optional" if repeated_num <= 1 else "repeated"
            self._LayoutOneField(field.rule, struct_name, field.name)

            self._col += (repeated_num-1) * (col_end-col_begin)

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
        self._output.append("* @file:   " + self._pb_file_name + "\n")
        self._output.append("* @author: jameyli <jameyli AT tencent DOT com>\n")
        self._output.append("* @brief:  这个文件是通过工具自动生成的，建议不要手动修改\n")
        self._output.append("*/\n")
        self._output.append("\n")


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

    def _LayoutOneField(self, field_rule, field_type, field_name) :
        """输出一行定义"""
        if not self._is_layout :
            return
        if field_name.find('=') > 0 :
            name_and_value = field_name.split('=')
            self._output.append(" "*self._indentation + field_rule + " " + field_type \
                    + " " + str(name_and_value[0]).strip() + " = " + self._GetAndAddFieldIndex()\
                    + " [default = " + str(name_and_value[1]).strip() + "]" + ";\n")
            return

        if (field_rule != "required" and field_rule != "optional") :
            self._output.append(" "*self._indentation + field_rule + " " + field_type \
                    + " " + field_name + " = " + self._GetAndAddFieldIndex() + ";\n")
            return

        if field_type == "bool" :
                self._output.append(" "*self._indentation + field_rule + " " + field_type \
                        + " " + field_name + " = " + self._GetAndAddFieldIndex()\
                        + " [default = false]" + ";\n")
        elif field_type == "int32" or field_type == "int64"\
                or field_type == "uint32" or field_type == "uint64"\
                or field_type == "sint32" or field_type == "sint64"\
                or field_type == "fixed32" or field_type == "fixed64"\
                or field_type == "sfixed32" or field_type == "sfixed64" \
                or field_type == "double" or field_type == "float" :
                    self._output.append(" "*self._indentation + field_rule + " " + field_type \
                            + " " + field_name + " = " + self._GetAndAddFieldIndex()\
                            + " [default = 0]" + ";\n")
        elif field_type == "string" or field_type == "bytes" :
            self._output.append(" "*self._indentation + field_rule + " " + field_type \
                    + " " + field_name + " = " + self._GetAndAddFieldIndex()\
                    + " [default = \"\"]" + ";\n")
        elif field_type == "DateTime" :
            field_type = "uint64"
            self._output.append(" "*self._indentation + field_rule + " " + field_type \
                    + " /*DateTime*/ " + field_name + " = " + self._GetAndAddFieldIndex()\
                    + " [default = 0]" + ";\n")
        elif field_type == "TimeDuration" :
            field_type = "uint64"
            self._output.append(" "*self._indentation + field_rule + " " + field_type \
                    + " /*TimeDuration*/ " + field_name + " = " + self._GetAndAddFieldIndex()\
                    + " [default = 0]" + ";\n")
        else :
            self._output.append(" "*self._indentation + field_rule + " " + field_type \
                    + " " + field_name + " = " + self._GetAndAddFieldIndex() + ";\n")
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
        file_path = PROTO_OUTPUT_PATH + self._pb_file_name
        #  os.makedirs(PROTO_OUTPUT_PATH)
        pb_file = open(file_path, "w+")
        pb_file.writelines(self._output)
        pb_file.close()

class DataParser:
    """解析excel的数据"""
    def __init__(self, xls_file_path, sheet):
        self._xls_file_path = xls_file_path
        self._sheet = sheet
        self._sheet_name = self._sheet.name

        self._row_count = len(self._sheet.col_values(0))
        self._col_count = len(self._sheet.row_values(0))

        self._row = 0
        self._col = 0

        try:
            self._module_name = OUTPUT_FULE_PATH_BASE + self._sheet_name.lower() + "_pb2"
            sys.path.append(os.getcwd())
            exec('from '+self._module_name + ' import *');
            self._module = sys.modules[self._module_name]
        except BaseException, e :
            print "load module(%s) failed"%(self._module_name)
            raise

    def Parse(self) :
        """对外的接口:解析数据"""
        LOG_INFO("begin parse, row_count = %d, col_count = %d", self._row_count, self._col_count)

        item_array = getattr(self._module, self._sheet_name+'_ARRAY')()

        # 先找到定义ID的列
        id_col = 0
        for id_col in range(self._col_count) :
            info_id = str(self._sheet.cell_value(self._row, id_col)).strip()
            if info_id == "" :
                continue
            else :
                break

        for self._row in range(DATA_BEGIN_ROW, self._row_count) :
            # 如果 id 是 空 直接跳过改行
            info_id = str(self._sheet.cell_value(self._row, id_col)).strip()
            if info_id == "" :
                LOG_WARN("%d is None", self._row)
                continue
            item = item_array.items.add()
            self._ParseLine(item)

        LOG_INFO("parse result:\n%s", item_array)

        self._WriteReadableData2File(str(item_array))

        data = item_array.SerializeToString()
        self._WriteData2File(data)

        LogHelp.close()

    def _ParseLine(self, item) :
        LOG_INFO("%d", self._row)

        self._col = 0
        while self._col < self._col_count :
            self._ParseField(item)

    def _ParseField(self, item) :
        field_rule = str(self._sheet.cell_value(FIELD_RULE_ROW, self._col)).strip()

        if field_rule == "required" or field_rule == "optional" :
            field_name = str(self._sheet.cell_value(FIELD_NAME_ROW, self._col)).strip()
            if field_name.find('=') > 0 :
                name_and_value = field_name.split('=')
                field_name = str(name_and_value[0]).strip()
            field_type = str(self._sheet.cell_value(FIELD_TYPE_ROW, self._col)).strip()

            LOG_INFO("%d|%d", self._row, self._col)
            LOG_INFO("%s|%s|%s", field_rule, field_type, field_name)

            field_value = self._GetFieldValue(field_type, self._row, self._col)
            # 有value才设值
            if field_value != None :
                item.__setattr__(field_name, field_value)
            self._col += 1

        elif field_rule == "repeated" :
            second_row = str(self._sheet.cell_value(FIELD_TYPE_ROW, self._col)).strip()
            LOG_DEBUG("repeated|%s", second_row);
            # 一般是简单的单字段，数值用分号相隔
            # 一般也只能是数字类型
            field_type = second_row
            field_name = str(self._sheet.cell_value(FIELD_NAME_ROW, self._col)).strip()
            field_value_str = unicode(self._sheet.cell_value(self._row, self._col))
            #增加长度判断
            if len(field_value_str) > 0:
                if field_value_str.find(";\n") > 0 :
                    field_value_list = field_value_str.split(";\n")
                else :
                    field_value_list = field_value_str.split(";")

                for field_value in field_value_list :
                    if len(field_value_str) <= 0:
                        break;

                    if field_type == "bytes" or field_type == "string" :
                        item.__getattribute__(field_name).append(field_value.encode("utf8"))
                    elif field_type == "double" or field_type == "float" :
                        item.__getattribute__(field_name).append(float(field_value))
                    else:
                        item.__getattribute__(field_name).append(int(float(field_value)))

            self._col += 1

        elif "struct" in field_rule :
            field_num = int(self._sheet.cell_value(FIELD_TYPE_ROW, self._col).split('*')[0])
            repeated_num = int(self._sheet.cell_value(FIELD_TYPE_ROW, self._col).split('*')[1])

            field_name = str(self._sheet.cell_value(FIELD_NAME_ROW, self._col)).strip()
            struct_name = "InternalType_" + field_name;

            LOG_INFO("%s|%d|%s|%s", field_rule, field_num, struct_name, field_name)


            self._col += 1

            # 至少循环一次
            if repeated_num <= 1 :
                struct_item = item.__getattribute__(field_name)
                self._ParseStruct(field_num, struct_item)
            else :
                for i in range(repeated_num):
                    struct_item = item.__getattribute__(field_name).add()
                    self._ParseStruct(field_num, struct_item)
                    if len(struct_item.ListFields()) == 0 :
                        # 空数据就不用加了
                        item.__getattribute__(field_name).__delitem__(-1)
        else :
            self._col += 1
            return

    def _ParseStruct(self, field_num, struct_item) :
        """嵌套结构数据读取"""

        # 跳过结构体定义
        for i in range(field_num) :
            self._ParseField(struct_item)

    def _GetFieldValue(self, field_type, row, col) :
        """将pb类型转换为python类型"""

        field_value = self._sheet.cell_value(row, col)
        LOG_INFO("%d|%d|%s", row, col, field_value)

        try:
            if field_type == "bool" :
                if len(str(field_value).strip()) <=0 :
                    return False
                else :
                    return (True if int(field_value) != 0 else False)
            elif field_type == "int32" or field_type == "int64"\
                    or  field_type == "uint32" or field_type == "uint64"\
                    or field_type == "sint32" or field_type == "sint64"\
                    or field_type == "fixed32" or field_type == "fixed64"\
                    or field_type == "sfixed32" or field_type == "sfixed64" :
                        if len(str(field_value).strip()) <=0 :
                            return None
                        else :
                            return int(field_value)
            elif field_type == "double" or field_type == "float" :
                    if len(str(field_value).strip()) <=0 :
                        return None
                    else :
                        return float(field_value)
            elif field_type == "string" :
                field_value = unicode(field_value)
                if len(field_value) <= 0 :
                    return None
                else :
                    return field_value
            elif field_type == "bytes" :
                field_value = unicode(field_value).encode('utf-8')
                if len(field_value) <= 0 :
                    return None
                else :
                    return field_value
            elif field_type == "DateTime" :
                field_value = unicode(field_value).encode('utf-8')
                if len(field_value) <= 0 :
                    return 0
                else :
                    import time
                    time_struct = time.strptime(field_value, "%Y-%m-%d %H:%M:%S")
                    #GMT time
                    timt_stamp = int(time.mktime(time_struct) - time.timezone)
                    return timt_stamp
            elif field_type == "TimeDuration" :
                field_value = unicode(field_value).encode('utf-8')
                if len(field_value) <= 0 :
                    return 0
                else :
                    import datetime
                    import time
                    time_struct=0
                    try :
                        time_struct = time.strptime(field_value, "%HH")
                    except BaseException, e :
                        time_struct = time.strptime(field_value, "%jD%HH")
                    return 3600 * (time_struct.tm_yday * 24 + time_struct.tm_hour)
            else :
                return None
        except BaseException, e :
            print "parse cell(%u, %u) error, please check it, maybe type is wrong."%(row, col)
            raise

    def _WriteData2File(self, data) :
        self._data_file_name = OUTPUT_FULE_PATH_BASE + self._sheet_name.lower() + ".data"
        file = open(self._data_file_name, 'wb+')
        file.write(data)
        file.close()

    def _WriteReadableData2File(self, data) :
        file_name = OUTPUT_FULE_PATH_BASE + self._sheet_name.lower() + ".txt"
        file = open(file_name, 'wb+')
        file.write(data)
        file.close()


###############################################################################
def ProcessOneFile(xls_file_path, op) :
    if not ".xls" in xls_file_path :
        #  print "Skip %s" %(xls_file_path)
        return

    try :
        workbook = xlrd.open_workbook(xls_file_path)
    except BaseException, e :
        print Color.RED + "Open %s failed! File is NOT exist!" %(xls_file_path) + Color.NONE
        sys.exit(-2)

    sheet_name_list = workbook.sheet_names()

    for sheet_name in sheet_name_list :
        if (not sheet_name.isupper()):
            print "Skip %s" %(sheet_name)
            continue

        try :
            sheet = workbook.sheet_by_name(sheet_name)
        except BaseException, e :
            print Color.YELLOW, e, Color.NONE
            break

        try :
            interpreter = SheetInterpreter(xls_file_path, sheet)
            interpreter.Interpreter()
        except BaseException, e :
            print Color.RED + "Interpreter %s of %s Failed!!!" %(sheet_name, xls_file_path) + Color.NONE
            print Color.YELLOW, "row: ", interpreter._row + 1, "col: ", interpreter._col + 1, "ERROR: ", e, Color.NONE
            break

        print Color.GREEN + "Interpreter %s of %s TO %s Success!!!" %(sheet_name, xls_file_path, interpreter._pb_file_name) + Color.NONE

        try :
            parser = DataParser(xls_file_path, sheet)
            parser.Parse()
        except BaseException, e :
            print Color.RED + "Parse %s of %s Failed!!!" %(sheet_name, xls_file_path) + Color.NONE
            print Color.YELLOW, "row: ", parser._row + 1, "col: ", parser._col + 1, "ERROR: ", e, Color.NONE
            break

        print Color.GREEN + "Parse %s of %s TO %s Success!!!" %(sheet_name, xls_file_path, parser._data_file_name) + Color.NONE

    workbook.release_resources()

def ProcessPath(file_path, op) :
    if os.path.isfile(file_path) :
        ProcessOneFile(file_path, op)
    elif os.path.isdir(file_path) :
        path_list = os.listdir(file_path)
        for path in path_list :
            real_file_path = file_path+"/"+path
            ProcessPath(real_file_path, op)

def usage():
    print '''
    Usage: %s EXCEL_FILE [OPTIONS]
    ''' %(sys.argv[0])


if __name__ == '__main__' :
    if len(sys.argv) < 2 :
        Usage()
        sys.exit(-1)

    #  try:
    #      opt, args = getopt.getopt(sys.argv[3:], "hr:s:o:", ["help", "repeated=", "subkey=", "output="])
    #  except getopt.GetoptError, err:
    #      print "err:",(err)
    #      usage()
    #      sys.exit(-1)

    #  for op, value in opt:
    #      if op == "-h" :
    #          usage()
    #      elif op == "-r" or op == "--repeated":
    #          have_repeated_key = int(value)
    #      elif op == "-s" or op == "--subkey":
    #          sub_key_col = int(value)
    #      elif op == "-o" or op == "--output":
    #          output = value


    # option 0 生成proto和data; 1 只生成proto; 2 只生成data
    op = 0
    if len(sys.argv) > 2 :
        op = int(sys.argv[2])

    xls_file_path =  sys.argv[1]
    ProcessPath(xls_file_path, op)


