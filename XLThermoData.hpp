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

#ifndef XLTHERMO_DATA_XLTHERMO_DATA_HPP
#define XLTHERMO_DATA_XLTHERMO_DATA_HPP

#include <xll.h>

xll::AddIn xllCompData(
    // Return double, C++ name of function, Excel name.
    xll::Function(XLL_LPOPER, "compData", "XLTHERMO.COMPDATA")
        // Array of function arguments.
        .Arguments({ xll::Arg(XLL_LPOPER, "Name", "is the name of the component to get data for."),
                     xll::Arg(XLL_LPOPER, "Property", "is the name of the property to get.") })
        // Function Wizard help.
        .FunctionHelp("Get pure component properties for a given component.")
        // Function Wizard category.
        .Category("Engineering")
        // URL linked to `Help on this function`.
        .HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal"));

// WINAPI calling convention must be specified
xll::LPOPER WINAPI compData(xll::LPOPER name, xll::LPOPER property);




// =============================================================================
// XLL.TGAMMA
// =============================================================================
xll::AddIn xai_tgamma2(
    // Return double, C++ name of function, Excel name.
    xll::Function(XLL_DOUBLE, "xll_tgamma", "TROLDAL.TGAMMA2", false)
        // Array of function arguments.
        .Arguments({ xll::Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.") })
        // Function Wizard help.
        .FunctionHelp("Return the Gamma function value.")
        // Function Wizard category.
        .Category("MATH")
        // URL linked to `Help on this function`.
        .HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal"));

// WINAPI calling convention must be specified
_FPX* WINAPI xll_tgamma(double x);

// =============================================================================
// XLL.TGAMMA2
// =============================================================================
xll::AddIn xai_tgamma(
    // Return double, C++ name of function, Excel name.
    xll::Function(XLL_LPOPER, "xll_tgamma2", "XLL.TGAMMA")
        // Array of function arguments.
        .Arguments({ xll::Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.") })
        // Function Wizard help.
        .FunctionHelp("Return the Gamma function value.")
        // Function Wizard category.
        .Category("MATH")
        // URL linked to `Help on this function`.
        .HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal"));

// WINAPI calling convention must be specified
xll::LPOPER WINAPI xll_tgamma2(double x);

// =============================================================================
// XLTHERMO.CMD.LOADCOMPLIST
// =============================================================================
// Press Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
// See https://xlladdins.github.io/Excel4Macros/
xll::AddIn xlmLoadComponentList(xll::Macro("loadComponentList", "XLTHERMO.CMD.LOADCOMPLIST"));

// Macros must have `int WINAPI (*)(void)` signature.
int WINAPI loadComponentList();

#endif    // XLTHERMO_DATA_XLTHERMO_DATA_HPP
