add_library(
    kwinto-xll SHARED

    "kwXll.cpp"
    "kwUtils.cpp"

    "framework/framework.c"
    "framework/MemoryManager.cpp"
    "framework/MemoryPool.cpp"
)

 set_target_properties(
     kwinto-xll
     PROPERTIES
        OUTPUT_NAME "kwinto"
        SUFFIX ".xll"
 )

include(FetchContent)

# Link nlohmann/json
FetchContent_Declare(
    json
    URL https://github.com/nlohmann/json/releases/download/v3.11.2/json.tar.xz
)
FetchContent_MakeAvailable(json)


target_include_directories(
    kwinto-xll
    
    PUBLIC
        framework
)

target_link_libraries(
    kwinto-xll

    PRIVATE
        nlohmann_json::nlohmann_json
        ws2_32
#        wsock32
)
