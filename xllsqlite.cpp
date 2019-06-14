// xllsqlite.cpp - sqlite wrapper
#include <locale>
#include "xllsqlite.h"

using namespace xll;

AddIn xai_sqlite(
    Documentation(L"Sqlite3 wrapper")
);

#if 0
AddIn xai_sql_select(
    Function(XLL_CSTRING, L"?xll_sql_select", L"SQL.SELECT")
    .Arg(XLL_LPOPER, L"columns", L"is an array of columns to select. Default is '*'")
    .Arg(XLL_LPOPER, L"from", L"is a range of one or more tables to select from.")
    .Arg(XLL_LPOPER, L"where", L"is a range of conjunctions and disjuntions.")
    .Arg(XLL_LPOPER, L"order", L"specifies the ORDER BY clause.")
    .Arg(XLL_LPOPER, L"group", L"specifies the GROUP BY clause.")
    .FunctionHelp(L"Return a SQL SELECT statement.")
    .Category(L"SQL")
    .Documentation(L"Documentation.")
);
const wchar_t* xll_sql_select(LPOPER columns, LPOPER from, LPOPER where, LPOPER order, LPOPER group)
{
#pragma XLLEXPORT
    static wstring select;

    try {
        select = L"";
        wstring comma;
        
        comma = L"SELECT ";
        for (const auto& c : *columns) {
            ensure (c.isStr());
            select.append(comma);
            select.append(c.val.str + 1, c.val.str[0]);
            comma = L", ";
        }

        comma = L"\nFROM ";
        for (const auto& f : *from) {
            ensure(f.isStr());
            select.append(comma);
            select.append(f.val.str + 1, f.val.str[0]);
            comma = L", ";
        }

        select.append(L"\nWHERE (");
        wstring rs = L"";
        const auto& w = *where;
        for (int i = 0; i < w.rows(); ++i) {
            for (int j = 0; j < w.columns(); ++j) {
                ensure(w(i,j).isStr());
                select.append(comma);
                select.append(w(i,j).val.str + 1, w(i,j).val.str[0]);
                comma = L") and (";
            }
        }
        select.append(L")");

    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        return 0;
    }

    return select.c_str();
}
#endif

XLL_ENUM_DOC(SQLITE_OPEN_READONLY, SQLITE_OPEN_READONLY, L"SQLITE", L"Read only access.", L"Read only access.");
XLL_ENUM_DOC(SQLITE_OPEN_READWRITE, SQLITE_OPEN_READWRITE, L"SQLITE", L"Read and write access.", L"Read and write access.");
XLL_ENUM_DOC(SQLITE_OPEN_CREATE, SQLITE_OPEN_CREATE, L"SQLITE", L"Create database if it does not already exist.", L"Create database if it does not already exist.");
XLL_ENUM_DOC(SQLITE_OPEN_URI, SQLITE_OPEN_URI, L"SQLITE", L"Open using Universal Resource Identifier.", L"Open using Universal Resource Identifier.");
XLL_ENUM_DOC(SQLITE_OPEN_MEMORY, SQLITE_OPEN_MEMORY, L"SQLITE", L"Open data base in memory.", L"Open data base in memory.");
XLL_ENUM_DOC(SQLITE_OPEN_NOMUTEX, SQLITE_OPEN_NOMUTEX, L"SQLITE", L"Do not use mutal exclusing when accessing database.", L"Do not use mutal exclusing when accessing database.");
XLL_ENUM_DOC(SQLITE_OPEN_FULLMUTEX, SQLITE_OPEN_FULLMUTEX, L"SQLITE", L"Use mutal exclusing when accessing database.", L"Use mutal exclusing when accessing database.");

AddIn xai_sqlite_db(
    Function(XLL_HANDLE, L"?xll_sqlite_db", L"SQLITE.DB")
    .Arg(XLL_CSTRING4, L"file", L"is the name of the sqlite3 database to open.")
    .Arg(XLL_SHORT, L"flags", L"is an optional set of flags from the SQLITE_OPEN_* enumeration to use when opening the database. Default is SQLITE_OPEN_READONLY.")
    .Uncalced()
    .FunctionHelp(L"Return a handle to a sqlite3 database.")
    .Category(L"SQLITE")
    .Documentation(L"Call sqlite3_open_v2 with flags SQLITE_OPEN_READONLY.")
);
HANDLEX WINAPI xll_sqlite_db(const char* file, SHORT flags)
{
#pragma XLLEXPORT
    handlex h;

    try {
        if (flags == 0)
            flags = SQLITE_OPEN_READONLY;

        handle<sqlite::db> h_(new sqlite::db(file, flags));
        h = h_.get();
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return h;
}

// insert range into table
#define XLL_LPOPER4 L"P"
AddIn xai_sqlite_insert(
    Function(XLL_HANDLE, L"?xll_sqlite_insert", L"SQLITE.INSERT")
    .Arg(XLL_HANDLE, L"db", L"is a handle to a database returned by SQLITE.OPEN.")
    .Arg(XLL_CSTRING4, L"table", L"is the name of the table.")
    .Arg(XLL_LPOPER, L"range", L"is the range to insert.")
    .FunctionHelp(L"Return the database handle.")
    .Category(L"SQLITE")
    .Documentation(L".")
);
HANDLEX WINAPI xll_sqlite_insert(HANDLEX db, const char* table, const XLOPER* pr)
{
#pragma XLLEXPORT
    handlex h;

    try {
        std::string sql = "insert into ";
        sql.append(table);
        // walk through rows of pr...
        OPER o;
        o = Excel(xlfEvaluate, *pr);
        db = db;
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
    .Arg(XLL_BOOL, L"?headers", L"is an optional argument to specify if headers should be included. Default is false.")
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
        ensure (h_.ptr());
        o = sqlite_range(*h_, sql, headers);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return &o;
}

AddIn xai_sqlite_tables(
    Function(XLL_LPOPER, L"?xll_sqlite_tables", L"SQLITE.TABLES")
    .Arg(XLL_HANDLE, L"handle", L"is the sqlite3 database handle returned by SQLITE.DB.")
    .FunctionHelp(L"Return a list of all tables in the database.")
    .Category(L"SQLITE")
    .Documentation(L"")
);
LPOPER WINAPI xll_sqlite_tables(handlex h)
{
#pragma XLLEXPORT
    static OPER o;

    try {
        static const char* sql = "SELECT name FROM sqlite_master WHERE type = \"table\"";
        handle<sqlite::db> h_(h);
        ensure(h_.ptr());
        o = sqlite_range(*h_, sql, false);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }


    return &o;
}

AddIn xai_sqlite_table_info(
    Function(XLL_LPOPER, L"?xll_sqlite_table_info", L"SQLITE.TABLE.INFO")
    .Arg(XLL_HANDLE, L"handle", L"is the sqlite3 database handle returned by SQLITE.DB.")
    .Arg(XLL_CSTRING4, L"table", L"is the name of a table in the database.")
    .Arg(XLL_BOOL, L"?headers", L"is an optional argument to specify if headers should be included. Default is false.")
    .FunctionHelp(L"Return the cid, name, type, default value, and whether it is a primary key.")
    .Category(L"SQLITE")
    .Documentation(L"")
);
LPOPER WINAPI xll_sqlite_table_info(handlex h, const char* table, BOOL headers)
{
#pragma XLLEXPORT
    static OPER o;

    try {
        std::string sql = "PRAGMA table_info(";
        sql.append(table);
        sql.append(")");
        handle<sqlite::db> h_(h);
        ensure(h_.ptr());
        o = sqlite_range(*h_, sql.c_str(), headers);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return &o;
}


#ifdef _DEBUG

xll::test test_sqlite_range([]{
    char dir[1024];
    GetCurrentDirectoryA(1023, dir);
    const char* file = "C:/Users/kal/Source/Repos/keithalewis/xllsqlite/chinook.db";
    const char* sql = "PRAGMA table_info(artists);";
    sqlite::db db(file);
    sqlite::db::stmt stmt(db);
    OPER o;
    o = sqlite_range(db, sql, true);
    o = o;

});

#endif // _DEBUG