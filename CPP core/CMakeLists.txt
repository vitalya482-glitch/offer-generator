cmake_minimum_required(VERSION 3.20)

project(CalcMaker LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

file(GLOB CALC_MAKER_SOURCES CONFIGURE_DEPENDS
    "${CMAKE_CURRENT_SOURCE_DIR}/CalcMakerClasses/*.cpp"
)

add_executable(CalcMaker
    CalcMakerMain.cpp
    ${CALC_MAKER_SOURCES}
)

target_include_directories(CalcMaker PRIVATE
    "${CMAKE_CURRENT_SOURCE_DIR}"
    "${CMAKE_CURRENT_SOURCE_DIR}/CalcMakerClasses"
)

if(MINGW)
    target_link_options(CalcMaker PRIVATE
        -static
        -static-libgcc
        -static-libstdc++
    )
endif()

set_target_properties(CalcMaker PROPERTIES
    OUTPUT_NAME "CalcMaker"
)
