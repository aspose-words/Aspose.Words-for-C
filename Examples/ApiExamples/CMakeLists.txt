cmake_minimum_required(VERSION 3.16.0 FATAL_ERROR)

project(ApiExamples CXX)

find_package(CodePorting.Native.Cs2Cpp REQUIRED CONFIG PATHS ${CMAKE_CURRENT_SOURCE_DIR}/../.. NO_DEFAULT_PATH)
find_package(Aspose.Words.Cpp REQUIRED CONFIG PATHS ${CMAKE_CURRENT_SOURCE_DIR}/../.. NO_DEFAULT_PATH)
find_package(Threads REQUIRED)

file(GLOB_RECURSE SRC_FILES "source/*.cpp" "source/*.h")
add_executable(ApiExamples ${SRC_FILES})

target_include_directories(ApiExamples PRIVATE source)
target_link_libraries(ApiExamples PRIVATE Aspose::Words Threads::Threads) 

add_subdirectory(../gtest gtest)
target_link_libraries(ApiExamples PRIVATE gtest)

target_compile_options(ApiExamples PRIVATE $<$<CXX_COMPILER_ID:MSVC>:/wd4244 /wd4250 /wd4267 /MP /bigobj /utf-8>)

set(CMAKE_MODULE_PATH ${CMAKE_CURRENT_SOURCE_DIR}/cmake)
include(GroupSources) 
GroupSources(${PROJECT_NAME} "${CMAKE_CURRENT_SOURCE_DIR}")


enable_testing()

file(TO_NATIVE_PATH "${Aspose.Words.Cpp_DIR}/lib/${CMAKE_VS_PLATFORM_NAME}" Aspose.Words.Cpp_DLL_PATH)
file(TO_NATIVE_PATH "${CodePorting.Native.Cs2Cpp_DIR}/lib" CodePorting.Native.Cs2Cpp_DLL_PATH)
file(TO_NATIVE_PATH "${CMAKE_CURRENT_SOURCE_DIR}" VS_WOKRING_DIRECTORY)

add_test(NAME ApiExamples COMMAND ApiExamples WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}")
if (WIN32)
    set_property(TEST ApiExamples PROPERTY ENVIRONMENT "PATH=%PATH%\;${Aspose.Words.Cpp_DLL_PATH}\\$<CONFIG>\;${CodePorting.Native.Cs2Cpp_DLL_PATH}")
endif()

set_directory_properties(PROPERTIES VS_STARTUP_PROJECT ApiExamples)

set_target_properties(ApiExamples PROPERTIES 
# Setup Visual Studio Project Properties
    VS_DEBUGGER_ENVIRONMENT "PATH=${Aspose.Words.Cpp_DLL_PATH}\\$<CONFIG>;${CodePorting.Native.Cs2Cpp_DLL_PATH};$(Path)"
    VS_DEBUGGER_WORKING_DIRECTORY "${VS_WOKRING_DIRECTORY}"
# Setup Unity Build
    UNITY_BUILD TRUE
# Setup required C++ standard
    CXX_STANDARD 14
    CXX_STANDARD_REQUIRED TRUE
)