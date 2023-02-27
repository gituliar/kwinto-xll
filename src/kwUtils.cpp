#include "kwUtils.h"


Error
kw::toJson(const XLOPER& key_, const XLOPER& value_, json& object)
{
    std::string key = "key";
    switch (key_.xltype)
    {
    case xltypeStr:
    {
        const auto& p = key_.val.str;
        key = std::string(p + 1, p[0]);
        break;
    }
    };

    switch (value_.xltype)
    {
    case xltypeNum:
        object[key] = value_.val.num;
        break;
    case xltypeStr:
    {
        const auto& p = value_.val.str;
        object[key] = std::string(p + 1, p[0]);
        break;
    }
    case xltypeMulti:
    {
        auto& array = object[key];
        const auto& array_ = value_.val.array;
        for (auto i = 0; i < array_.rows; ++i)
        {
            for (auto j = 0; j < array_.columns; ++j)
            {
                auto& cell = array_.lparray[i * array_.columns + j];
                switch (cell.xltype)
                {
                case xltypeNum:
                    array.push_back(cell.val.num);
                    break;
                case xltypeStr:
                    const auto& p = cell.val.str;
                    array.push_back(std::string(p + 1, p[0]));
                    break;
                }
            }
        }
        break;
    }
    case xltypeMissing:
        object[key] = nullptr;
        break;
    };

    return "";
}