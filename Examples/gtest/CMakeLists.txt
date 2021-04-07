cmake_minimum_required(VERSION 3.7 FATAL_ERROR)

project(gtest CXX)

add_library(gtest STATIC gtest/gtest.h gtest/gtest-spi.h gmock/gmock.h gmock_main.cc gmock-gtest-all.cc)

find_package(Threads)
target_link_libraries(gtest PUBLIC Threads::Threads)

target_compile_definitions(gtest PUBLIC GTEST_INCLUDES_AVAILABLE=1)
target_include_directories(gtest SYSTEM PUBLIC ${CMAKE_CURRENT_SOURCE_DIR})
