// xllsqlite.cpp - sqlite wrapper
#include <locale>
#include "xllsqlite.h"

using namespace xll;

AddIn xai_sqlite(
    Documentation(L"Sqlite3 wrapper")
);

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
    sqlite::db db(file);
    sqlite::db::stmt stmt(db);
    OPER o;
    o = sqlite_range(db, sql, true);
    o = o;

});

#endif // _DEBUG