# CMakeList.txt : CMake project for kwinto-xll, include source and define
# project specific logic here.
#
cmake_minimum_required (VERSION 3.8)

project ("kwinto-xll")

# Add source to this project's executable.
add_executable (kwinto-xll "kwinto-xll.cpp" "kwinto-xll.h")

if (CMAKE_VERSION VERSION_GREATER 3.12)
  set_property(TARGET kwinto-xll PROPERTY CXX_STANDARD 20)
endif()

# TODO: Add tests and install targets if needed.
