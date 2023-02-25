add_library(
    kwinto-xll SHARED

    "kwXll.cpp"

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

target_include_directories(
    kwinto-xll
    
    PUBLIC
        framework
)

target_link_libraries(
    kwinto-xll

    PRIVATE
        ws2_32
#        wsock32
)
