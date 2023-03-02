#pragma once

#include "kwCommon.h"

namespace kw
{

enum xlOper
{
    Empty  = 0b0000'0001,
    Number = 0b0000'0010,
    String = 0b0000'0100,
    Vector = 0b0000'1000,
    Matrix = 0b0001'0000,
    Any    = Empty | Number | String | Vector | Matrix
};

struct RangeAttributes
{
    char
        axis; // x | y
    size_t
        size; // number of cells along |axis|
    size_t
        type;

    RangeAttributes() :
        axis { 'y' }, size{ 0 }, type { xlOper::Any }
    {};
};

Error
    toJson(const XLOPER& oper, json& object, RangeAttributes& attr);
Error
    toJson(const XLOPER& key_, const XLOPER& value_, json& object);

}