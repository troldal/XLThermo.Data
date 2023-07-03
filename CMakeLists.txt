#==============================================================================
# Preamble
#==============================================================================
cmake_minimum_required(VERSION 3.24)
project(XLThermo.Data)

#==============================================================================
# Set compiler settings
#==============================================================================
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
set(CMAKE_CXX_VISIBILITY_PRESET hidden)
set(CMAKE_C_VISIBILITY_PRESET hidden)
set(CMAKE_VISIBILITY_INLINES_HIDDEN 1)

#==============================================================================
# Add the dependency (xll and PCData) libraries to the project
#==============================================================================
add_subdirectory(deps)

#==============================================================================
# Add XLThermo.Data library and set source file paths
#==============================================================================
add_library(
        XLThermoData
        SHARED
        ${CMAKE_CURRENT_SOURCE_DIR}/XLThermoData.cpp
)

#==============================================================================
# Set Xll library properties and compile options
#==============================================================================
target_compile_options(XLThermoData PRIVATE "/J" "/utf-8" )
target_link_libraries(XLThermoData PRIVATE xll::xlcall xll::xll)
set_target_properties(XLThermoData PROPERTIES OUTPUT_NAME "XLThermo.Data" SUFFIX ".xll")
target_include_directories(XLThermoData PUBLIC ${CMAKE_CURRENT_BINARY_DIR})
