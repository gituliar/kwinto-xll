#include "kwUtils.h"



Error
kw::toJson(const XLOPER& value_, json& value, RangeAttributes& attr)
{
    if ((attr.type & xlOper::Empty) && (value_.xltype == xltypeMissing))
    {
        value = nullptr;
        return "";
    }

    if ((attr.type & xlOper::Number) && (value_.xltype == xltypeNum))
    {
        value = value_.val.num;
        return "";
    };

    if ((attr.type & xlOper::String) && (value_.xltype == xltypeStr))
    {
        const auto& p = value_.val.str;
        value = std::string(p + 1, p[0]);
        return "";
    };

    if ((attr.type & xlOper::Vector) && (value_.xltype == xltypeMulti))
    {
        const auto& array_ = value_.val.array;
        if (((array_.rows == 1) && ((attr.size == 0 && array_.columns > 1) || (attr.size == array_.columns))) ||
            ((array_.columns == 1) && ((attr.size == 0 && array_.rows > 1) || (attr.size == array_.rows))))
        {
            for (auto i = 0; i < array_.rows; ++i)
            {
                for (auto j = 0; j < array_.columns; ++j)
                {
                    const auto& cell_ = array_.lparray[i * array_.columns + j];

                    RangeAttributes cellAttr;
                    cellAttr.type = xlOper::Number | xlOper::Number | xlOper::String;

                    json cell;
                    if (auto error = kw::toJson(cell_, cell, cellAttr); !error.empty())
                        return "" + error;

                    value.push_back(cell);
                }
            }

            attr.axis = array_.rows > 1 ? 'y' : 'x';
            attr.size = value.size();

            return "";
        }
    }

    if ((attr.type & xlOper::Matrix) && (value_.xltype == xltypeMulti))
    {
        const auto& array_ = value_.val.array;

        auto size = (attr.axis == 'x') ? array_.columns : array_.rows;
        if (attr.size == 0 || attr.size == size)
        {
            if (attr.axis == 'x')
                attr.size = array_.columns;
            else
            {
                attr.axis = 'y';
                attr.size = array_.rows;
            }

            RangeAttributes cellAttr;
            cellAttr.type = xlOper::Empty | xlOper::Number | xlOper::String;

            for (auto i = 0; i < array_.rows; ++i)
            {
                for (auto j = 0; j < array_.columns; ++j)
                {
                    auto& array = (attr.axis == 'y') ? value[i] : value[j];
                    const auto& cell_ = array_.lparray[i * array_.columns + j];

                    json value;
                    if (auto error = kw::toJson(cell_, value, cellAttr); !error.empty())
                        return error;

                    array.push_back(value);
                }
            }

            return "";
        };
    }

    return "kw::toJson: Fail to parse Excel range to JSON";
};


Error
kw::toJson(const XLOPER& key_, const XLOPER& value_, json& value)
{
    // Valid <key, value> combinations:
    //   - empty,  string
    //   - string, empty
    //   - string, number
    //   - string, string
    //   - string, vector
    //   - string, matrix
    //   - vector, vector
    //   - vector, matrix

    RangeAttributes keyAttr;
    keyAttr.type = xlOper::Empty | xlOper::String | xlOper::Vector;

    json key;
    if (auto error = toJson(key_, key, keyAttr); !error.empty())
        return error;


    if (key.empty())
    {
        RangeAttributes valueAttr;
        valueAttr.type = xlOper::String;

        if (auto error = toJson(value_, value[key], valueAttr); !error.empty())
            return error;

        return "";
    }
    else if (key.is_string())
    {
        RangeAttributes valueAttr;
        valueAttr.type = xlOper::Empty | xlOper::Number | xlOper::String | xlOper::Vector | xlOper::Matrix;

        if (auto error = toJson(value_, value[key.get<std::string>()], valueAttr); !error.empty())
            return error;

        return "";
    }
    else if (key.is_array())
    {
        for (const auto& name : key.array())
        {
            if (!name.is_string())
                return "kw::toJson: JSON key should be string";
        };

        RangeAttributes valueAttr;
        valueAttr.axis = keyAttr.axis;
        valueAttr.size = key.size();
        valueAttr.type = xlOper::Vector | xlOper::Matrix;

        json matrix;
        if (auto error = toJson(value_, matrix, valueAttr); !error.empty())
            return error;

        for (auto i = 0; i < key.size(); ++i)
            value[key[i]] = matrix[i];

        return "";
    }

    return "kw::toJson: Fail to parse Excel range to JSON";
}