cmake_minimum_required(VERSION 3.5)

project(DocumentExplorer LANGUAGES CXX)

set(CMAKE_INCLUDE_CURRENT_DIR ON)

set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(QT NAMES Qt6 Qt5 COMPONENTS Widgets REQUIRED)
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Widgets REQUIRED PATHS ${QT_DIR})

find_package(CodePorting.Native.Cs2Cpp REQUIRED)
find_package(Aspose.Words.Cpp REQUIRED)

set(PROJECT_SOURCES
    main.cpp
    mainwindow.cpp
    mainwindow.h
    mainwindow.ui
    viewerwindow.cpp
    viewerwindow.h
    viewerwindow.ui
    aboutdialog.cpp
    aboutdialog.h
    aboutdialog.ui
    documentmodel.cpp
    documentmodel.h
    hidpisupport.cpp
    hidpisupport.h
    qtcorehelpers.h
    resources.qrc
)

if(${QT_VERSION_MAJOR} GREATER_EQUAL 6)
    qt_add_executable(${PROJECT_NAME} ${PROJECT_SOURCES})
else()
    add_executable(${PROJECT_NAME} ${PROJECT_SOURCES})
endif()

set_property(TARGET ${PROJECT_NAME} PROPERTY WIN32_EXECUTABLE true)

target_link_libraries(${PROJECT_NAME} PRIVATE Qt${QT_VERSION_MAJOR}::Widgets Aspose::Words)

add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different $<TARGET_FILE:Aspose::CodePorting.Native.Cs2Cpp> $<TARGET_FILE_DIR:${PROJECT_NAME}>
    COMMAND ${CMAKE_COMMAND} -E copy_if_different $<TARGET_FILE:Aspose::Words> $<TARGET_FILE_DIR:${PROJECT_NAME}>
)

if(WIN32)
    get_target_property(QMAKE_EXECUTABLE Qt${QT_VERSION_MAJOR}::qmake IMPORTED_LOCATION)
    get_filename_component(WINDEPLOYQT_EXECUTABLE ${QMAKE_EXECUTABLE} PATH)
    set(WINDEPLOYQT_EXECUTABLE "${WINDEPLOYQT_EXECUTABLE}/windeployqt.exe")

    add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
       COMMAND ${WINDEPLOYQT_EXECUTABLE} --no-compiler-runtime $<TARGET_FILE:${PROJECT_NAME}>)
endif()
