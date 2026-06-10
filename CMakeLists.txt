cmake_minimum_required(VERSION 3.18)
project(xlsxtext VERSION 1.0.0 LANGUAGES C CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# --- Library target (header-only) ---
add_library(xlsxtext INTERFACE)
target_include_directories(xlsxtext INTERFACE
    ${CMAKE_CURRENT_SOURCE_DIR}/include
)
target_sources(xlsxtext INTERFACE
    ${CMAKE_CURRENT_SOURCE_DIR}/include/xlsxtext.hpp
    ${CMAKE_CURRENT_SOURCE_DIR}/include/number_format.hpp
    ${CMAKE_CURRENT_SOURCE_DIR}/include/miniz/miniz.c
    ${CMAKE_CURRENT_SOURCE_DIR}/include/miniz/miniz.h
    ${CMAKE_CURRENT_SOURCE_DIR}/include/pugixml/pugixml.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/include/pugixml/pugixml.hpp
    ${CMAKE_CURRENT_SOURCE_DIR}/include/pugixml/pugiconfig.hpp
)

# --- Demo / integration test target ---
add_executable(number_format_test test/number_format.test.cpp)
target_link_libraries(number_format_test PRIVATE xlsxtext)

add_executable(xlsx_test test/xlsx.test.cpp)
target_link_libraries(xlsx_test PRIVATE xlsxtext)