cmake_minimum_required(VERSION 3.15)
project(Hakaton)

set(CMAKE_CXX_STANDARD 17)

find_package(OpenXLSX REQUIRED)

find_package(SQLite3 REQUIRED)

add_executable(Hakaton main.cpp)

target_link_libraries(Hakaton PRIVATE OpenXLSX::OpenXLSX SQLite::SQLite3)