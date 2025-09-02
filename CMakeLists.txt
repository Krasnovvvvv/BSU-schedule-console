cmake_minimum_required(VERSION 3.28)
project(BSU-schedule)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(fmt CONFIG REQUIRED)
find_package(CURL CONFIG REQUIRED)
find_package(xlsxio CONFIG REQUIRED)

add_executable(BSU-schedule main.cpp)

target_link_libraries(BSU-schedule PRIVATE
        fmt::fmt
        CURL::libcurl
        xlsxio::xlsxio_read_SHARED xlsxio::xlsxio_write_SHARED
)