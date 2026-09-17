cmake_minimum_required(VERSION 3.21 FATAL_ERROR)

project(DocsExamples CXX)

find_package(aspose-words-cpp CONFIG REQUIRED)
find_package(Threads REQUIRED)

# Aspose.Email for C++ is an optional second product, used only by the
# BaseConversions.DocxToMhtml example. It ships as a NuGet package
# (Aspose.Email.Cpp) and is wired up by DocsExamples.vcxproj on Windows.
# When it is not present, that one example is compiled out.
set(ASPOSE_EMAIL_ROOT "${CMAKE_CURRENT_SOURCE_DIR}/../../packages/Aspose.Email.Cpp.26.8.0/build/native"
    CACHE PATH "Root of the Aspose.Email for C++ NuGet package")

# Aspose.BarCode for C++ is an optional second product, used only by the
# WorkingWithBarcodeGenerator example. It ships as a NuGet package
# (Aspose.BarCode.Cpp) and is wired up by DocsExamples.vcxproj on Windows.
# When it is not present, that one example is compiled out.
set(ASPOSE_BARCODE_ROOT "${CMAKE_CURRENT_SOURCE_DIR}/../../packages/Aspose.BarCode.Cpp.26.8.0/build/native"
    CACHE PATH "Root of the Aspose.BarCode for C++ NuGet package")

file(GLOB_RECURSE SRC_FILES "source/*.cpp" "source/*.h")
add_executable(DocsExamples ${SRC_FILES})

target_include_directories(DocsExamples PRIVATE source)
target_link_libraries(DocsExamples PRIVATE Aspose::Words Threads::Threads) 

if(WIN32 AND EXISTS "${ASPOSE_EMAIL_ROOT}/include/Aspose.Email.Cpp")
    target_compile_definitions(DocsExamples PRIVATE ASPOSE_EMAIL_AVAILABLE)
    target_include_directories(DocsExamples PRIVATE "${ASPOSE_EMAIL_ROOT}/include/Aspose.Email.Cpp")
    if(CMAKE_SIZEOF_VOID_P EQUAL 8)
        set(_ae_arch x64)
        set(_ae_tag vc14x64)
    else()
        set(_ae_arch x86)
        set(_ae_tag vc14x86)
    endif()
    target_link_libraries(DocsExamples PRIVATE
        "$<$<CONFIG:Debug>:${ASPOSE_EMAIL_ROOT}/lib/${_ae_arch}/Debug/Aspose.Email.Cpp_${_ae_tag}d.lib>"
        "$<$<NOT:$<CONFIG:Debug>>:${ASPOSE_EMAIL_ROOT}/lib/${_ae_arch}/Release/Aspose.Email.Cpp_${_ae_tag}.lib>")
    add_custom_command(TARGET DocsExamples POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<IF:$<CONFIG:Debug>,${ASPOSE_EMAIL_ROOT}/lib/${_ae_arch}/Debug/Aspose.Email.Cpp_${_ae_tag}d.dll,${ASPOSE_EMAIL_ROOT}/lib/${_ae_arch}/Release/Aspose.Email.Cpp_${_ae_tag}.dll>"
            $<TARGET_FILE_DIR:DocsExamples>)
else()
    message(STATUS "Aspose.Email for C++ not found - BaseConversions.DocxToMhtml will be skipped.")
endif()

if(WIN32 AND EXISTS "${ASPOSE_BARCODE_ROOT}/include")
    target_compile_definitions(DocsExamples PRIVATE ASPOSE_BARCODE_AVAILABLE)
    target_include_directories(DocsExamples PRIVATE "${ASPOSE_BARCODE_ROOT}/include")
    if(CMAKE_SIZEOF_VOID_P EQUAL 8)
        set(_ab_tag vc14x64)
    else()
        set(_ab_tag vc14x86)
    endif()
    target_link_libraries(DocsExamples PRIVATE
        "$<$<CONFIG:Debug>:${ASPOSE_BARCODE_ROOT}/lib/Aspose.BarCode.Cpp_${_ab_tag}d.lib>"
        "$<$<NOT:$<CONFIG:Debug>>:${ASPOSE_BARCODE_ROOT}/lib/Aspose.BarCode.Cpp_${_ab_tag}.lib>")
    add_custom_command(TARGET DocsExamples POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<IF:$<CONFIG:Debug>,${ASPOSE_BARCODE_ROOT}/lib/Aspose.BarCode.Cpp_${_ab_tag}d.dll,${ASPOSE_BARCODE_ROOT}/lib/Aspose.BarCode.Cpp_${_ab_tag}.dll>"
            $<TARGET_FILE_DIR:DocsExamples>)
else()
    message(STATUS "Aspose.BarCode for C++ not found - WorkingWithBarcodeGenerator will be skipped.")
endif()

add_subdirectory(../gtest gtest)
target_link_libraries(DocsExamples PRIVATE gtest)

add_subdirectory(../SQLiteCpp SQLiteCpp)
target_link_libraries(DocsExamples PRIVATE SQLiteCpp)

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