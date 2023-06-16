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

// xll_template.cpp - Sample xll project.
#include <cmath>    // for double tgamma(double)
#include <xll.h>

using namespace xll;

AddIn xai_tgamma(
    // Return double, C++ name of function, Excel name.
    Function(XLL_DOUBLE, "xll_tgamma", "TROLDAL.TGAMMA")
        // Array of function arguments.
        .Arguments({ Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.") })
        // Function Wizard help.
        .FunctionHelp("Return the Gamma function value.")
        // Function Wizard category.
        .Category("MATH")
        // URL linked to `Help on this function`.
        .HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
        .Documentation(R"xyzyx(
The <i>Gamma</i> function is \(\Gamma(x) = \int_0^\infty t^{x - 1} e^{-t}\,dt\), \(x \ge 0\).
If \(n\) is a natural number then \(\Gamma(n + 1) = n! = n(n - 1)\cdots 1\).
<p>
Any valid HTML using <a href="https://katex.org/" target="_blank">KaTeX</a> can 
be used for documentation.
)xyzyx"));
// WINAPI calling convention must be specified
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT    // must be specified to export function

    return tgamma(x);
}

// Press Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
// See https://xlladdins.github.io/Excel4Macros/
AddIn xai_macro(
    // C++ function, Excel name of macro
    Macro("xll_macro", "XLL.MACRO"));
// Macros must have `int WINAPI (*)(void)` signature.
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
    // https://xlladdins.github.io/Excel4Macros/reftext.html
    // A1 style instead of default R1C1.
    OPER reftext = Excel(xlfReftext, Excel(xlfActiveCell), OPER(true));
    // UTF-8 strings can be used.
    Excel(xlcAlert, OPER("XLL.MACRO called with активный 细胞: ") & reftext);

    return TRUE;
}

#ifdef _DEBUG

// Create XML documentation and index.html in `$(TargetPath)` folder.
// Use `xsltproc file.xml -o file.html` to create HTML documentation.
Auto<Open> xao_template_docs([]() {
    return Documentation("MATH", "Documentation for the xll_template add-in.");
});

#endif    // _DEBUG
