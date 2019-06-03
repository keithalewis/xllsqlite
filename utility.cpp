#include "xll12/xll/xll.h"

using namespace xll;

AddIn xai_get_name(
    Function(XLL_LPOPER, L"?xll_get_name", L"GET_NAME")
    .FunctionHelp(L"Returns the full pathname of the add-in.")
    .Category(L"XLL")
    //    .Documentation(L"")
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

AddIn xai_open_dialog(
    Function(XLL_LPOPER, L"?xll_file_open_dialog", L"FILE.OPEN.DIALOG")
    .Arg(XLL_LPOPER, L"filter", L"is the file filtering criteria.")
    .Uncalced()
    .FunctionHelp(L"Open dialog to select file name.")
    .Category(L"XLL")
    .Documentation(L"")
);
LPOPER WINAPI xll_file_open_dialog(LPOPER pfilter)
{
#pragma XLLEXPORT
    static OPER o;
    static OPER default_filter = OPER(L"All Files (*.*)");
    static OPER default_button = OPER(L"Open");
    static OPER default_title = OPER(L"Open File");
    static OPER default_index = OPER(1);

    if (pfilter->isMissing()) {
        pfilter = &default_filter;
    }
    o = Excel(xlfOpenDialog, *pfilter, default_button, default_title, default_index);

    return &o;
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
