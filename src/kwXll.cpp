// Taken from https://github.com/elnormous/HTTPRequest
#include "extern/HTTPRequest.hpp"
#include "framework/framework.h"
#include "nlohmann/json.hpp"

using namespace std::string_literals;

using json = nlohmann::json;
using Error = std::string;


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


extern "C" __declspec(dllexport)
LPXLOPER
kwJson(const char* name, const char* key1_, const char* value1_)
{
    const std::string key(key1_);
    const std::string value(value1_);

    // create JSON object
    json obj = {
        {key, value}
    };

    // save JSON object
    std::string newName;
    if (auto error = JsonHub::instance().put(name, obj, newName); !error.empty())
    {
        return TempStrConst((LPSTR)("ERROR: kwJson: "s + error).c_str());
    }

    return TempStrConst((LPSTR)newName.c_str());
}


extern "C" __declspec(dllexport)
LPXLOPER
kwJsonShow(const char* name)
{
    json object;
    if (auto error = JsonHub::instance().get(name, object); !error.empty())
    {
        return TempStrConst((LPSTR)("ERROR: kwJsonShow: "s + error).c_str());
    }

    return TempStrConst((LPSTR)(object.dump()).c_str());
}


extern "C" __declspec(dllexport)
LPXLOPER
kwRPC(const char* method_, const char* name_, const char* jsonId_)
{
    const std::string method(method_);
    const std::string name(name_);

    const json params = {
        {"json", jsonId_}
    };


    json response;
    if (auto error = RpcClient::instance().send(method_, params, response); !error.empty())
        return TempStrConst((LPSTR)("ERROR: kwRPC: "s + error).c_str());

    if (response.contains("error"))
    {
        const auto& error = response["error"];

        if (!error.contains("message") || !error["message"].is_string())
        {
            auto buf = error.dump();
            buf.resize(255);

            return TempStrConst((LPSTR)("ERROR: kwRPC: Invalid error format: "s + buf).c_str());
        }

        return TempStrConst((LPSTR)("ERROR: kwRPC: "s + error["message"].get<std::string>()).c_str());
    }
    if (!response.contains("result"))
    {
        auto buf = response.dump();
        buf.resize(255);

        return TempStrConst((LPSTR)("ERROR: kwRPC: Invalid response format: "s + buf).c_str());
    }

    // put response to JsonHub
    std::string newName;
    if (auto error = JsonHub::instance().put(name, response["result"], newName); !error.empty())
        return TempStrConst((LPSTR)("ERROR: kwRPC: "s + error).c_str());


    return TempStrConst((LPSTR)newName.c_str());
}


extern "C" __declspec(dllexport)
int
xlAutoOpen(void)
{
    XLOPER12 xDLL;

    Excel12f(xlGetName, &xDLL, 0);

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"kwJson"),
        (LPXLOPER12)TempStr12(L"PCCC"),
        (LPXLOPER12)TempStr12(L"kwJson"),
        (LPXLOPER12)TempStr12(L"name,key1,value1"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"kwintoFunction"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Create a new JSON object in the local memory"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"kwJsonShow"),
        (LPXLOPER12)TempStr12(L"PC"),
        (LPXLOPER12)TempStr12(L"kwJsonShow"),
        (LPXLOPER12)TempStr12(L"name"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"kwintoFunction"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Show full JSON object"),
        (LPXLOPER12)TempStr12(L""));

    Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
        (LPXLOPER12)TempStr12(L"kwRPC"),
        (LPXLOPER12)TempStr12(L"PCCC"),
        (LPXLOPER12)TempStr12(L"kwRPC"),
        (LPXLOPER12)TempStr12(L"Method,Output Json Id,Input Json Id"),
        (LPXLOPER12)TempStr12(L"1"),
        (LPXLOPER12)TempStr12(L"kwintoFunction"),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L""),
        (LPXLOPER12)TempStr12(L"Call remote procedure on the remote host"),
        (LPXLOPER12)TempStr12(L""));


    // Free the XLL filename
    Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

    return 1;
}
