#pragma once

#include <tr1/unordered_map>
#include <vector>
#include <string>
#include <utility>

#include <google/protobuf/text_format.h>

#include "commlib/logging.h"
#include "commlib/random_util.h"

using namespace std;

template<typename INFO, typename INFO_ARRAY, int KEY_NUM = 1>
class DeployDataMgr {
public:
    static const unsigned int MAX_DATA_BUFF_LEN = 1024 * 1024 * 4;

    class Key {
        public:
            Key() {
                key_[0] = 0;
                key_[1] = 0;
                key_[2] = 0;
                key_[3] = 0;
            }

            explicit Key(unsigned int key1, unsigned int key2 = 0, unsigned int key3 = 0, unsigned int key4 = 0) {
                key_[0] = key1;
                key_[1] = key2;
                key_[2] = key3;
                key_[3] = key4;
            }

            unsigned int key_[4];
    };

    class HashOfKey {
        public:
            size_t operator() (const Key& key) const {
                size_t real_key = key.key_[0];
                return (real_key * (1ull << 32) + key.key_[1]);
            }
    };

    class EqualOfKey {
        public:
            bool operator() (const Key& rhs, const Key& lhs) const {
                return (rhs.key_[0] == lhs.key_[0] && rhs.key_[1] == lhs.key_[1]
                        && rhs.key_[2] == lhs.key_[2] && rhs.key_[3] == lhs.key_[3]);
            }
    };


protected:
    DeployDataMgr() {}
    virtual ~DeployDataMgr() {}

public:
    static DeployDataMgr& Instance() {
        static DeployDataMgr* instance = NULL;
        if (NULL == instance) {
            instance = new DeployDataMgr();
        }

        return *instance;
    }

 public:
    /**
     * @brief:  从PB二进制文件初始化
     *
     * @param  data_file pb二进制数据文件
     *
     * @return: 0 成功，
     *          非0 失败
     */
    int InitFromFile(const char* data_file) {
        assert(KEY_NUM <= 4);
        assert(data_file);
        if (!data_file) {
            LOG_ERROR(0, 0, "data_file is NULL");
            return -1;
        }

        _info_map.clear();

        LOG_DEBUG(0, 0, "data_file is %s", data_file);

        FILE* file = fopen(data_file, "rb");
        assert(file);
        if (!file) {
            LOG_ERROR(0, 0, "open file(%s) error", data_file);
            return -2;
        }

        char data[MAX_DATA_BUFF_LEN];
        size_t readn = fread(data, 1, sizeof(data), file);
        LOG_DEBUG(0, 0, "read size(%lu) from file(%s)", readn, data_file);
        fclose(file);

        bool parse_ret = _info_array.ParseFromArray(data, readn);
        if (!parse_ret) {
            LOG_ERROR(0, 0, "%s|ParseFromString failed", data_file);
            return -3;
        }
        LOG_DEBUG(0, 0, "ParseFromString succ");

        LOG_INFO(0, 0, "%s|count|%u", _info_array.GetTypeName().c_str(), _info_array.items_size());
        LOG_INFO(0, 0, "%s|detail|%s", _info_array.GetTypeName().c_str(), _info_array.ShortDebugString().c_str());

        if (_info_array.items_size() > 0) {
            default_info.CopyFrom(_info_array.items(0));
        }

        for (int i=0; i < _info_array.items_size(); ++i) {
            INFO* info = _info_array.mutable_items(i);
            if (NULL == info) {
                LOG_ERROR(0, 0, "%s|info is NULL index = %u", data_file, i);
                continue;
            }

            if(KEY_NUM > 0){
				const google::protobuf::Descriptor* desc = info->GetDescriptor();
				if (NULL == desc) {
					continue;
				}

				const google::protobuf::Reflection* ref = info->GetReflection();
				if (NULL == ref) {
					continue;
				}

				Key key;
				for (int k = 0; k < KEY_NUM; ++k) {
					// 默认第一个元素是索引
					const google::protobuf::FieldDescriptor* key_desc = desc->FindFieldByNumber(k+1);
					if (NULL == key_desc) {
						continue;
					}

					if (key_desc->cpp_type() == key_desc->CPPTYPE_INT32) {
						key.key_[k] = ref->GetInt32(*info, key_desc);
					} else if (key_desc->cpp_type() == key_desc->CPPTYPE_UINT32) {
						key.key_[k] = ref->GetUInt32(*info, key_desc);
					} else {
						key.key_[k] = 0;
					}
				}


				std::pair<InfoMapIter, bool> result =
					_info_map.insert(InfoMapValueType(key, info));
				if (!result.second) {
					LOG_ERROR(0, 0, "%s|insert info failed, index=%u, key = (%u, %u, %u, %u)",
							data_file, i, key.key_[0], key.key_[1], key.key_[2], key.key_[3]);
					continue;
				}

				LOG_INFO(0, 0, "%s|insert info succes, index=%u, key = (%u, %u, %u, %u)",
						data_file, i, key.key_[0], key.key_[1], key.key_[2], key.key_[3]);
            }
            LOG_INFO(0, 0, "%s|%s", (*info).GetTypeName().c_str(), (*info).ShortDebugString().c_str());
        }

        return 0;
    }

    /**
     * @brief: 获取一条配置信息
     *
     * @param  info_id 配置信息的ID
     *
     * @return: 配置信息的指针
     *          没有找到则返回NULL
     */
    // 这里参数顺序要使用者自己保证
    const INFO* GetOneInfo(unsigned int key1, unsigned int key2 = 0,
                           unsigned int key3 = 0, unsigned int key4 = 0) const {

        Key key(key1, key2, key3, key4);
        InfoMapIter iter = _info_map.find(key);
        if (iter == _info_map.end()) {
            return NULL;
        } else {
            return iter->second;
        }
    }

    const INFO& GetDefault() {
        return default_info;
    }

    const INFO_ARRAY& GetAllInfo() const {
        return _info_array;
    }

    const INFO* GetRandomOne() const {
        if (_info_array.items_size() <=0) {
            return NULL;
        }
        int idx = RandomUtil::Random(_info_array.items_size());
        return &_info_array.items(idx);
    }

    size_t InfoCount() const {
        return _info_array.items_size();
    }

 private:
    static DeployDataMgr* _instance;

 private:
    typedef typename std::tr1::unordered_map<Key, INFO*, HashOfKey, EqualOfKey > InfoMap;
    typedef typename InfoMap::const_iterator InfoMapIter;
    typedef typename InfoMap::value_type InfoMapValueType;

    INFO_ARRAY _info_array;
    InfoMap _info_map;
    INFO default_info;
};

