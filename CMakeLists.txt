cmake_minimum_required(VERSION 3.18)
project(xlsxtext)

set(CMAKE_CXX_STANDARD 11)

include_directories("./include")

file(GLOB_RECURSE src "./include/*.c" "./include/*.cpp" "./src/*.cpp")
add_executable(${PROJECT_NAME} ${src})
