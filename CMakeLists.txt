cmake_minimum_required(VERSION 3.18)
project(xlsxtext VERSION 1.0.1 LANGUAGES C CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# --- Library target (static library with .cpp compilation units) ---
add_library(xlsxtext STATIC
    ${CMAKE_CURRENT_SOURCE_DIR}/xlsxtext/miniz/miniz.c
    ${CMAKE_CURRENT_SOURCE_DIR}/xlsxtext/pugixml/pugixml.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/xlsxtext/number_format.cpp
)
target_include_directories(xlsxtext PUBLIC
    ${CMAKE_CURRENT_SOURCE_DIR}/xlsxtext
)

# --- Demo / integration test target ---
add_executable(number_format_test test/number_format.test.cpp)
target_link_libraries(number_format_test PRIVATE xlsxtext)

add_executable(xlsx_test test/xlsx.test.cpp)
target_link_libraries(xlsx_test PRIVATE xlsxtext)