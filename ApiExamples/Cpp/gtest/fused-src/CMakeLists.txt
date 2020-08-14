cmake_minimum_required(VERSION 3.7 FATAL_ERROR)

project(gtest CXX C)

find_package(Threads)

file(GLOB_RECURSE gtest_files "gtest/*.h" "gtest/*.cc" "gtest/*.cpp")

add_library(gtest STATIC ${gtest_files})

target_link_libraries(gtest PRIVATE Aspose::BuildSettings PUBLIC Threads::Threads)

target_include_directories(gtest SYSTEM PUBLIC ${CMAKE_CURRENT_SOURCE_DIR})

# library output name
include(GuessCompilerPrefix)
GUESS_COMPILER_PREFIX(_SUFFIX)
set_target_properties(gtest PROPERTIES OUTPUT_NAME "gtest_${_SUFFIX}")
set_target_properties(gtest PROPERTIES DEBUG_POSTFIX "d")
get_target_property(_output_name gtest OUTPUT_NAME)
message(STATUS "Output name: '" ${_output_name} "'")
