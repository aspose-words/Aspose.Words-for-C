cmake_minimum_required(VERSION 3.21 FATAL_ERROR)

project(DocsExamples CXX)

find_package(aspose-words-cpp CONFIG REQUIRED)
find_package(Threads REQUIRED)

file(GLOB_RECURSE SRC_FILES "source/*.cpp" "source/*.h")
add_executable(DocsExamples ${SRC_FILES})

target_include_directories(DocsExamples PRIVATE source)
target_link_libraries(DocsExamples PRIVATE Aspose::Words Threads::Threads) 

add_subdirectory(../gtest gtest)
target_link_libraries(DocsExamples PRIVATE gtest)

target_compile_options(DocsExamples PRIVATE $<$<CXX_COMPILER_ID:MSVC>:/wd4244 /wd4250 /wd4267 /MP /bigobj /utf-8>)

set(CMAKE_MODULE_PATH ${CMAKE_CURRENT_SOURCE_DIR}/cmake)
include(GroupSources) 
GroupSources(${PROJECT_NAME} "${CMAKE_CURRENT_SOURCE_DIR}")

if(CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
    execute_process(COMMAND ${CMAKE_CXX_COMPILER} -fuse-ld=gold -Wl,--version ERROR_QUIET OUTPUT_VARIABLE LD_VERSION)
    if ("${LD_VERSION}" MATCHES "GNU gold")
        target_link_options(DocsExamples PRIVATE -fuse-ld=gold -Wl,--fatal-warnings)
    endif()
elseif(CMAKE_CXX_COMPILER_ID STREQUAL "Clang")
    execute_process(COMMAND ${CMAKE_CXX_COMPILER} -fuse-ld=lld -Wl,--version ERROR_QUIET OUTPUT_VARIABLE LD_VERSION)
    if ("${LD_VERSION}" MATCHES "LLD")
        target_link_options(DocsExamples PRIVATE -fuse-ld=lld -Wl,--fatal-warnings)
    endif()
endif()

enable_testing()

file(TO_NATIVE_PATH "${CMAKE_CURRENT_SOURCE_DIR}" VS_WOKRING_DIRECTORY)

add_test(NAME DocsExamples COMMAND DocsExamples WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}")
if (WIN32)
  add_custom_command(TARGET DocsExamples POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different $<TARGET_RUNTIME_DLLS:DocsExamples> $<TARGET_FILE_DIR:DocsExamples>
    COMMAND_EXPAND_LISTS
  )
endif()

set_directory_properties(PROPERTIES VS_STARTUP_PROJECT DocsExamples)

set_target_properties(DocsExamples PROPERTIES 
# Setup Visual Studio Project Properties
    VS_DEBUGGER_WORKING_DIRECTORY "${VS_WOKRING_DIRECTORY}"
# Setup Unity Build
    UNITY_BUILD TRUE
# Setup required C++ standard
    CXX_STANDARD 14
    CXX_STANDARD_REQUIRED TRUE
)