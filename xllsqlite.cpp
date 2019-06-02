// xllsqlite.cpp - sqlite wrapper
#include <locale>
#include "xllsqlite.h"

using namespace xll;

AddIn xai_sqlite(
    Documentation(L"Sqlite3 wrapper")
);

AddIn xai_get_name(
    Function(XLL_LPOPER, L"?xll_get_name", L"GET_NAME")
//    .Uncalced()
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

AddIn xai_sqlite_db(
    Function(XLL_HANDLE, L"?xll_sqlite_db", L"SQLITE.DB")
    .Arg(XLL_CSTRING4, L"file", L"is the name of the sqlite3 database to open.")
    .Uncalced()
    .FunctionHelp(L"Return a handle to a sqlite3 database.")
    .Category(L"SQLITE")
    .Documentation(L"")
);
HANDLEX WINAPI xll_sqlite_db(const char* file)
{
#pragma XLLEXPORT
    handlex h;

    try {
        handle<sqlite::db> h_(new sqlite::db(file));
        h = h_.get();
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return h;
}

AddIn xai_sqlite_exec(
    Function(XLL_LPOPER, L"?xll_sqlite_exec", L"SQLITE.EXEC")
    .Arg(XLL_HANDLE, L"handle", L"is the sqlite3 database handle returned by SQLITE.DB.")
    .Arg(XLL_CSTRING4, L"sql", L"is the SQL query to execute on the database.")
    .Arg(XLL_BOOL, L"?headers", L"is an optional argument to specify if headers should be included.")
    .FunctionHelp(L"Return the result of executing a SQL command on a database.")
    .Category(L"SQLITE")
    .Documentation(L"")
);
LPOPER WINAPI xll_sqlite_exec(HANDLEX h, const char* sql, BOOL headers)
{
#pragma XLLEXPORT
    static OPER o;

    try {
        handle<sqlite::db> h_(h);
        ensure (h_);
        o = sqlite_range(*h_, sql, headers);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return &o;
}

#ifdef _DEBUG

xll::test test_sqlite_range([]{
    const char* file = "C:\\Users\\kalx\\source\\repos\\keithalewis\\xllsqlite\\chinook.db";
    const char* sql = "SELECT * FROM artists";
    /*
    DWORD dwAttrib = GetFileAttributes(file);
    bool b;
    b = dwAttrib & FILE_ATTRIBUTE_DIRECTORY;
    */
    sqlite::db db(file);
    sqlite::db::stmt stmt(db);
    OPER o;
    //o = sqlite_range(db, L"select * from artists", true);
    o = sqlite_range(db, sql, true);
    o = o;

});

#endif // _DEBUG