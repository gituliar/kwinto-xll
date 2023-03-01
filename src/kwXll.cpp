#include "kwCommon.h"
#include "kwUtils.h"


class JsonHub
{
public:
    std::map<std::string, json>
        m_jsonMap;

public:
    static JsonHub&
        instance()
    {
        static JsonHub hub;
        return hub;
    }

    Error
        get(const std::string& name, json& object)
    {
        auto ii = m_jsonMap.find(name);
        if (ii == m_jsonMap.end())
            return "JsonHub::get: No JSON with name '" + name + "'";

        object = ii->second;

        return "";
    }

    Error
        put(const std::string& name, const json& object, std::string& newName)
    {
        m_jsonMap[name] = object;

        newName = name;

        return "";
    }

private:
    JsonHub() {};
};


class MemoryHub
{
private:
    std::string
        m_string;
    std::vector<std::string>
        m_stringPool;

public:
    std::string&
        string()
    {
        return m_string;

        m_stringPool.emplace_back();
        return *m_stringPool.rbegin();
    }
};


class RpcClient
{
private:
    size_t
        m_id;
    size_t
        m_timeout;
    std::string
        m_uri;

public:
    size_t&
        timeout() { return m_timeout; }
    std::string&
        uri() { return m_uri; }

    static RpcClient&
        instance()
    {
        static RpcClient client;
        return client;
    }

    Error
        send(const std::string& method, const json& params, json& result)
    {
        auto id = m_id++;

        http::Response response;
        try
        {
            json request = {
                {"id", id},
                {"method", method},
                {"params", params}
            };

            response = http::Request(m_uri).send(
                "POST",
                request.dump(),
                {
                    {"Connection", "close"},
                    {"Content-Type", "application/json-rpc"}
                },
                std::chrono::milliseconds{m_timeout}
            );
        }
        catch (const std::exception& err)
        {
            return "ERROR: RpcClient::send : #1 "s + err.what();
        };

        try
        {
            result = json::parse(response.body, nullptr, false);
            return "";
        }
        catch (const json::parse_error& err)
        {
            return "ERROR: RpcClient::send : #2 "s + err.what();
        };
    };

private:
    RpcClient() :
        m_id{ 0 }, m_timeout{ 2000 }, m_uri { "http://localhost:4000" }
    {};

};


static MemoryHub memory;


extern "C" __declspec(dllexport)
const char *
kwJson(const char* name, XLOPER key1_, XLOPER value1_)
{
    Error error;
    auto& res = memory.string();

    // create JSON object
    json obj;
    if (error = kw::toJson(key1_, value1_, obj); !error.empty())
        return  (res = "ERROR: kwJson: " + error).c_str();

    // save JSON object
    if (auto error = JsonHub::instance().put(name, obj, res); !error.empty())
        return  (res = "ERROR: kwJson: " + error).c_str();

    return res.c_str();
}


extern "C" __declspec(dllexport)
const char *
kwShow(const char* name)
{
    auto& res = memory.string();

    json object;
    if (auto error = JsonHub::instance().get(name, object); !error.empty())
    {
        res = "ERROR: kwShow: " + error;
        return res.c_str();
    }

    res = object.dump();

    return res.c_str();
}


extern "C" __declspec(dllexport)
LPXLOPER
kwRpc(const char* method_, const char* name_, const char* jsonId_)
{
    const std::string method(method_);
    const std::string name(name_);

    const json params = {
        {"json", jsonId_}
    };


    json response;
    if (auto error = RpcClient::instance().send(method_, params, response); !error.empty())
        return TempStrConst((LPSTR)("ERROR: kwRpc: "s + error).c_str());

    if (response.contains("error"))
    {
        const auto& error = response["error"];

        if (!error.contains("message") || !error["message"].is_string())
        {
            auto buf = error.dump();
            buf.resize(255);

            return TempStrConst((LPSTR)("ERROR: kwRpc: Invalid error format: "s + buf).c_str());
        }

        return TempStrConst((LPSTR)("ERROR: kwRpc: "s + error["message"].get<std::string>()).c_str());
    }
    if (!response.contains("result"))
    {
        auto buf = response.dump();
        buf.resize(255);

        return TempStrConst((LPSTR)("ERROR: kwRpc: Invalid response format: "s + buf).c_str());
    }

    // put response to JsonHub
    std::string newName;
    if (auto error = JsonHub::instance().put(name, response["result"], newName); !error.empty())
        return TempStrConst((LPSTR)("ERROR: kwRpc: "s + error).c_str());


    return TempStrConst((LPSTR)newName.c_str());
}


extern "C" __declspec(dllexport)
const char *
kwValue(const char* id, const char* key_)
{
    auto& res = memory.string();

    json object;
    if (auto error = JsonHub::instance().get(id, object); !error.empty())
        return (res = "ERROR: kwValue: " + error).c_str();


    std::string key(key_);
    if (!object.contains(key))
        return (res = "ERROR: kwValue: Key '" + key + "' not found in JSON '" + id + "'").c_str();

    res = object[key].dump();

    return res.c_str();
}


extern "C" __declspec(dllexport)
int
xlAutoOpen(void)
{
    XLOPER12 xDLL;

    Excel12f(xlGetName, &xDLL, 0);

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"kwJson"),
        (LPXLOPER12)TempStr12(L"CCPP"),
        (LPXLOPER12)TempStr12(L"kwJson"),
        (LPXLOPER12)TempStr12(L"name,key1,value1"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"kwintoFunction"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Create a new JSON object in the local memory"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"kwShow"),
        (LPXLOPER12)TempStr12(L"CC"),
        (LPXLOPER12)TempStr12(L"kwShow"),
        (LPXLOPER12)TempStr12(L"name"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"kwintoFunction"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Show JSON object"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"kwRpc"),
        (LPXLOPER12)TempStr12(L"PCCC"),
        (LPXLOPER12)TempStr12(L"kwRpc"),
        (LPXLOPER12)TempStr12(L"Method,Output Json Id,Input Json Id"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"kwintoFunction"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Call remote procedure on the remote host"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"kwValue"),
        (LPXLOPER12)TempStr12(L"CCC"),
        (LPXLOPER12)TempStr12(L"kwValue"),
        (LPXLOPER12)TempStr12(L"Id,Key"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"kwintoFunction"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Show Json value"),
        (LPXLOPER12)TempStr12(L""));


    // Free the XLL filename
    Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

    return 1;
}
