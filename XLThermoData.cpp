/*

Y88b   d88P 888    88888888888 888
 Y88b d88P  888        888     888
  Y88o88P   888        888     888
   Y888P    888        888     88888b.   .d88b.  888d888 88888b.d88b.   .d88b.
   d888b    888        888     888 "88b d8P  Y8b 888P"   888 "888 "88b d88""88b
  d88888b   888        888     888  888 88888888 888     888  888  888 888  888
 d88P Y88b  888        888     888  888 Y8b.     888     888  888  888 Y88..88P
d88P   Y88b 88888888   888     888  888  "Y8888  888     888  888  888  "Y88P"

MIT License

Copyright (c) 2023 Kenneth Troldal Balslev

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

*/

#include "XLThermoData.hpp"
#include <iostream>
#include <tuple>
#include <vector>
#include <xll.h>
#include <PCData.hpp>
#include <nlohmann/json.hpp>

using namespace xll;

namespace
{

    /**
     * @brief
     * @param errCode
     * @return
     */
    xll::LPOPER makeError(int errCode)
    {
        auto* result    = new OPER();
        result->xltype  = xltypeErr | xlbitDLLFree;
        result->val.err = errCode;
        return result;
    }
}    // namespace

xll::LPOPER compData(xll::LPOPER name, xll::LPOPER property)
{
#pragma XLLEXPORT

    // Ensure that the name argument is a string. If missing, return #N/A
    if (name->is_missing()) return makeError(xlerrNA);
    if (!name->is_str()) return makeError(xlerrValue);

    // Ensure that the property argument is either missing or is a string.
    if (!property->is_missing() && !property->is_str()) return makeError(xlerrValue);

    // Convert the name and property arguments to local std::string objects
    auto theName     = name->to_string();
    auto theProperty = (property->is_str() ? property->to_string() : "");

    // Free the memory allocated by Excel for the name and property arguments
    ::Excel12(xlFree, nullptr, 1, name);
    ::Excel12(xlFree, nullptr, 1, property);

    static const PCData db("./pcd.db");
    nlohmann::json comp = nlohmann::json::parse(db.find(theName));

    if (theProperty.empty()) {
        auto* rng   = new OPER(3, 1);
        rng->xltype = xltypeMulti | xlbitDLLFree;
        (*rng)[0]   = OPER(comp["MolecularWeight"]["Value"].get<double>());
        (*rng)[1]   = OPER(comp["CriticalTemperature"]["Value"].get<double>());
        (*rng)[2]   = OPER(comp["CriticalPressure"]["Value"].get<double>());
        return rng;
    }

    if (theProperty == "MW") {
        auto* result = new OPER(3.14);
        result->xltype |= xlbitDLLFree;
        return result;
    }

    if (theProperty == "Tc") {
        auto* result = new OPER(2.71);
        result->xltype |= xlbitDLLFree;
        return result;
    }

    if (theProperty == "Pc") {
        auto* result = new OPER(1.41);
        result->xltype |= xlbitDLLFree;
        return result;
    }
}

// WINAPI calling convention must be specified
_FPX* WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT    // must be specified to export function

    static xll::FPX arr(2, 2);
    arr(0, 0) = 3.14;
    arr(0, 1) = 2.71;
    arr(1, 0) = 1.41;
    arr(1, 1) = 1.61;

    return arr.get();
}

LPOPER WINAPI xll_tgamma2(double x)
{
#pragma XLLEXPORT    // must be specified to export function

    auto components = std::vector< std::tuple< std::string, std::string, double > > {};
    components.emplace_back("Water", "H2O", 1.0);
    components.emplace_back("Ethanol", "C2H6O", 2.0);
    components.emplace_back("Methane", "CH4", 3.0);
    components.emplace_back("Propane", "C3H8", 4.0);

    auto* rng   = new OPER(components.size(), 3);
    rng->xltype = xltypeMulti | xlbitDLLFree;
    for (size_t i = 0; i < components.size(); ++i) {
        (*rng)(i, 0) = std::get< 0 >(components[i]).c_str();
        (*rng)(i, 1) = std::get< 1 >(components[i]).c_str();
        (*rng)(i, 2) = std::get< 2 >(components[i]);
    }

    //        auto* rng   = new OPER(2, 2);
    //        rng->xltype = xltypeMulti | xlbitDLLFree;
    //        (*rng)[0]   = OPER("Blah");
    //        (*rng)[1]   = OPER(2.0);
    //        (*rng)[2]   = OPER(3.0);
    //        (*rng)[3]   = OPER(4.0);

    std::cerr << "xll_tgamma2 called with x = " << x << std::endl;

    return rng;
}

int WINAPI loadComponentList()
{
#pragma XLLEXPORT

    static auto rng = OPER(xltypeMulti);
    rng.xltype      = xltypeMulti | xlbitDLLFree;
    rng.resize(2, 2);
    rng(0, 0) = "Blah";
    rng(0, 1) = 2.0;
    rng(1, 0) = 3.0;
    rng(1, 1) = 4.0;

    auto ref                 = Excel(xlfActiveCell);
    ref.val.sref.ref.colLast = ref.val.sref.ref.colFirst + 1;
    ref.val.sref.ref.rwLast  = ref.val.sref.ref.rwFirst + 1;
    Excel(xlSet, ref, rng);

    return TRUE;
}
