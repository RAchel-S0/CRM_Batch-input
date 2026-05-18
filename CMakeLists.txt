cmake_minimum_required(VERSION 3.16)
project(CRM_Automation VERSION 3.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)

find_package(Qt6 REQUIRED COMPONENTS Widgets Network)

add_executable(CRM_Automation WIN32
    src/main.cpp
    src/webdriver_client.h
    src/webdriver_client.cpp
    src/logic_handler.h
    src/logic_handler.cpp
    src/main_window.h
    src/main_window.cpp
)

target_link_libraries(CRM_Automation PRIVATE
    Qt6::Widgets
    Qt6::Network
)

configure_file(
    ${CMAKE_SOURCE_DIR}/reply_list.txt
    ${CMAKE_BINARY_DIR}/reply_list.txt
    COPYONLY
)
