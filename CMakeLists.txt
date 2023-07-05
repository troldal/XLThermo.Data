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
# Set the path to the PCData library
#==============================================================================
set(FETCHCONTENT_SOURCE_DIR_PCDATA ${CMAKE_CURRENT_LIST_DIR}/../PCData)
include(FetchContent)
FetchContent_Declare(pcdata
        GIT_REPOSITORY https://github.com/troldal/PCData.git
        GIT_TAG        main
        )
FetchContent_MakeAvailable(pcdata)

#==============================================================================
# Import JSON library
#==============================================================================
FetchContent_Declare(json
        GIT_REPOSITORY https://github.com/nlohmann/json.git
        GIT_TAG bc889af
        )
FetchContent_MakeAvailable(json)

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
        Database.cpp Database.hpp)

#==============================================================================
# Set Xll library properties and compile options
#==============================================================================
target_compile_options(XLThermoData PRIVATE "/J" "/utf-8" )
target_link_libraries(XLThermoData PRIVATE xll::xlcall xll::xll PCData::PCData nlohmann_json::nlohmann_json)
set_target_properties(XLThermoData PROPERTIES OUTPUT_NAME "XLThermo.Data" SUFFIX ".xll")
target_include_directories(XLThermoData PUBLIC ${CMAKE_CURRENT_BINARY_DIR})
file(COPY_FILE ./database/ChemSep8.32.xml ${CMAKE_CURRENT_BINARY_DIR}/pcd.db)
