// xllsqlite.cpp - sqlite wrapper
#include <locale>
#include "xllsqlite.h"

using namespace xll;

AddIn xai_sqlite(
    Documentation(L"Sqlite3 wrapper")
);

AddIn xai_sqlite_db(
    Function(XLL_HANDLE, L"?xll_sqlite_db", L"SQLITE.DB")
    .Arg(XLL_CSTRING, L"file", L"is the name of the sqlite3 database to open.")
    .Uncalced()
    .FunctionHelp(L"Return a handle to a sqlite3 database.")
    .Category(L"SQLITE")
    .Documentation(L"")
);
HANDLEX WINAPI xll_sqlite_db(const XCHAR* file)
{
#pragma XLLEXPORT
    handlex h;

    try {
        char buf[1024];
        size_t n;
        wcstombs_s(&n, buf, file, 1024);
//        handle<sqlite::db> h_(new sqlite::db(buf));
        handle<sqlite::db> h_(new sqlite::db(file));
        ensure (h_);
        h = h_.get();
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return h;
}