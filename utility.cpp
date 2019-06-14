#include "xll12/xll/xll.h"
#include <commdlg.h>


using wstring = std::basic_string<wchar_t>;
using namespace xll;

AddIn xai_get_name(
    Function(XLL_LPOPER, L"?xll_get_name", L"GET_NAME")
    .FunctionHelp(L"Returns the full pathname of the add-in.")
    .Category(L"XLL")
    .Documentation(L"")
);
LPOPER WINAPI xll_get_name()
{
#pragma XLLEXPORT
    static OPER o;

    o = Args::XlGetName();

    return &o;
}

AddIn xai_get_workspace(
    Function(XLL_LPOPER, L"?xll_get_workspace", L"GET_WORKSPACE")
    .Arg(XLL_SHORT, L"type_num", L"is a number specifying the type of workspace information you want. .")
    .Uncalced()
    .FunctionHelp(L"Returns information about the workspace.")
    .Category(L"XLL")
    .Documentation(L"")
);
LPOPER WINAPI xll_get_workspace(SHORT type_num)
{
#pragma XLLEXPORT
    static OPER o;

    o = Excel(xlfGetWorkspace, OPER(type_num));

    return &o;
}

AddIn xai_open_file(
    Function(XLL_CSTRING, L"?xll_open_file", L"OPEN.FILE")
    .FunctionHelp(L"Display a dialog for selecting a file name.")
    .Category(L"XLL")
    .Documentation(L"Return the name of the selected file.")
);
const wchar_t* WINAPI xll_open_file()
{
#pragma XLLEXPORT
    static wchar_t buf[1024];
    OPENFILENAME ofn;

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.nMaxFile = 1023;
    ofn.lpstrFile = buf;
    buf[0] = 0;
    ofn.Flags = OFN_FILEMUSTEXIST|OFN_PATHMUSTEXIST;
    BOOL b;
    b = GetOpenFileName(&ofn);

    return buf;
}

AddIn xai_join(
    Function(XLL_CSTRING, L"?xll_join", L"JOIN")
    .Arg(XLL_LPOPER, L"range", L"is a two-dimensional range.")
    .Arg(XLL_CSTRING, L"fs", L"is the field seperator for rows.")
    .Arg(XLL_CSTRING, L"rs", L"is the record seperator for columns.")
    .Arg(XLL_CSTRING, L"bracket", L"is an optional two character string for the first and last characters.")
    .FunctionHelp(L"Join cells in range using separators and return a string.")
    .Category(L"XLL")
    .Documentation(L"")
);
const wchar_t* WINAPI xll_join(LPOPER range, xcstr fs, xcstr rs, xcstr bracket)
{
#pragma XLLEXPORT
    static wstring join;

    try {
        join = L"";
        if (bracket && bracket[0]) {
            join.append(1, bracket[0]);
        }
        wstring FS = L"";
        wstring RS = L"";
        const auto& r = *range;
        for (int i = 0; i < r.rows(); ++i) {
            join.append(RS);
            for (int j = 0; j < r.columns(); ++j) {
                join.append(FS);
                OPER rij = Excel(xlfText, r(i, j), OPER(L"General"));
                ensure(rij.isStr());
                join.append(rij.val.str + 1, rij.val.str[0]);
                FS = fs;
            }
            RS = rs;
        }
        if (bracket && bracket[0] && bracket[1]) {
            join.append(1, bracket[1]);
        }
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        return 0;
    }

    return join.c_str();
}

AddIn xai_directory(
    Function(XLL_LPOPER, L"?xll_directory", L"CURRENT.DIRECTORY")
    .Arg(XLL_LPOPER, L"directory", L"is the path of the directory.")
    .Uncalced()
    .FunctionHelp(L"Set the current directory or return current directory.")
    .Category(L"XLL")
    .Documentation(L"")
);
LPOPER WINAPI xll_directory(LPOPER pdirectory)
{
#pragma XLLEXPORT
    static OPER o;

    o = Excel(xlfDirectory, *pdirectory);

    return &o;
}
