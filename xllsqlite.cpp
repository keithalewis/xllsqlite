// xllsqlite.cpp - sqlite wrapper
#include <locale>
#include "xllsqlite.h"
#include <commdlg.h>

using namespace xll;

using wstring = std::basic_string<wchar_t>;

AddIn xai_sqlite(
    Documentation(L"Sqlite3 wrapper")
);

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
                OPER rij = Excel(xlfText, r(i,j), OPER(L"General"));
                ensure (rij.isStr());
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