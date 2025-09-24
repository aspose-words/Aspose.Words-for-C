cmake_minimum_required(VERSION 3.21.0 FATAL_ERROR)
set(CMAKE_CXX_STANDARD 14)

project(ApiExamples CXX)

find_package(CodePorting.Translator.Cs2Cpp.Framework REQUIRED CONFIG)
find_package(Aspose.Words.Cpp REQUIRED CONFIG)
find_package(Aspose.Words.Shaping.HarfBuzz.Cpp REQUIRED CONFIG)

find_package(Threads REQUIRED)

file(GLOB_RECURSE SRC_FILES "source/*.cpp" "source/*.h")

set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${PROJECT_BINARY_DIR}/bin)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${PROJECT_BINARY_DIR}/bin)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${PROJECT_BINARY_DIR}/bin)

add_executable(ApiExamples ${SRC_FILES})

target_include_directories(ApiExamples PRIVATE source)
target_link_libraries(ApiExamples PRIVATE Aspose::Words Aspose::WordsShapingHarfBuzz Threads::Threads) 

add_subdirectory(../gtest gtest)
target_link_libraries(ApiExamples PRIVATE gtest)

target_compile_options(ApiExamples PRIVATE $<$<CXX_COMPILER_ID:MSVC>:/wd4244 /wd4250 /wd4267 /wd4996 /MP /bigobj /utf-8>)

set(CMAKE_MODULE_PATH ${CMAKE_CURRENT_SOURCE_DIR}/cmake)
include(GroupSources) 
GroupSources(${PROJECT_NAME} "${CMAKE_CURRENT_SOURCE_DIR}")

if(CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
    execute_process(COMMAND ${CMAKE_CXX_COMPILER} -fuse-ld=gold -Wl,--version ERROR_QUIET OUTPUT_VARIABLE LD_VERSION)
    if ("${LD_VERSION}" MATCHES "GNU gold")
        target_link_options(ApiExamples PRIVATE -fuse-ld=gold -Wl,--fatal-warnings)
    endif()
elseif(CMAKE_CXX_COMPILER_ID STREQUAL "Clang")
    execute_process(COMMAND ${CMAKE_CXX_COMPILER} -fuse-ld=lld -Wl,--version ERROR_QUIET OUTPUT_VARIABLE LD_VERSION)
    if ("${LD_VERSION}" MATCHES "LLD")
        target_link_options(ApiExamples PRIVATE -fuse-ld=lld -Wl,--fatal-warnings)
    endif()
endif()

enable_testing()

if (WIN32)
  add_custom_command(TARGET ApiExamples POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different $<TARGET_RUNTIME_DLLS:ApiExamples> $<TARGET_FILE_DIR:ApiExamples>
    COMMAND_EXPAND_LISTS
  )
endif()

add_test(NAME ApiExamples COMMAND ApiExamples WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}")

set_directory_properties(PROPERTIES VS_STARTUP_PROJECT ApiExamples)

file(TO_NATIVE_PATH "${CMAKE_CURRENT_SOURCE_DIR}" VS_WOKRING_DIRECTORY)
set_target_properties(ApiExamples PROPERTIES 
    VS_DEBUGGER_WORKING_DIRECTORY "${VS_WOKRING_DIRECTORY}"
    UNITY_BUILD TRUE
    CXX_STANDARD 14
    CXX_STANDARD_REQUIRED TRUE
)
