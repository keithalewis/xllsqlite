// xllsqlite.h - sqlite3 wrapper
#pragma once
#include "sqlite-amalgamation-3280000/sqlite3.h"
#include "xll12/xll/xll.h"

namespace sqlite {

    class value {
        sqlite3_value* val;
    public:
        value()
            : val(sqlite3_value_dup(0))
        { }
        value(const value& v)
            : val(sqlite3_value_dup(v.val))
        { }
        value& operator=(const value& v)
        {
            return *this = value(v);
        }
        value(value&& v)
            : val(v.val)
        {
            v.val = sqlite3_value_dup(0);
        }
        value& operator=(value&& v) noexcept
        {
            std::swap(val, v.val);

            return *this;
        }
        ~value()
        {
            sqlite3_value_free(val);
        }

        int bytes() const
        {
            return sqlite3_value_bytes(val);
        }
        int type() const
        {
            return sqlite3_value_type(val);
        }
    };

    // Sqlite converts wide strings to UTF-8 so we avoid *16* functions.
    class db {
        sqlite3* pdb;
    public:
        db(const char* file, int flags = SQLITE_OPEN_READONLY)
        {
            if (SQLITE_OK != sqlite3_open_v2(file, &pdb, flags, 0))
                throw std::runtime_error(sqlite3_errmsg(pdb));
        }
        db(const db&) = delete;
        db& operator=(const db&) = delete;
        ~db()
        {
            sqlite3_close(pdb);
        }
        // for use in sqlite3_* functions
        operator sqlite3*() {
            return pdb;
        }
        class stmt {
            sqlite::db& db;
            sqlite3_stmt* pstmt;
            const char* tail_;
        public:
            stmt(sqlite::db& db)
                : db(db), pstmt(nullptr)
            { }
            stmt(const stmt&) = delete;
            stmt& operator=(const stmt&) = delete;
            ~stmt()
            {
                if (pstmt != nullptr)
                    sqlite3_finalize(pstmt);
            }
            // for use in sqlite3_* functions
            operator sqlite3_stmt*()
            {
                return pstmt;
            }
            const char* errmsg() const
            {
                return sqlite3_errmsg(db);
            }
            int prepare(const char* sql, int nsql = -1)
            {
                return sqlite3_prepare_v2(db, sql, nsql, &pstmt, &tail_);
            }
            const char* tail() const
            {
                return tail_;
            }
            int bind(int col, int i)
            {
                return sqlite3_bind_int(pstmt, col, i);
            }
            int bind(int col, sqlite_int64 i)
            {
                return sqlite3_bind_int64(pstmt, col, i);
            }
            int bind(int col, double d)
            {
                return sqlite3_bind_double(pstmt, col, d);
            }
            // Do not make a copy of text by default.
            int bind(int col, const char* t, int n = -1, void(*dealloc)(void*) = SQLITE_STATIC)
            {
                return sqlite3_bind_text(pstmt, col, t, n, dealloc);
            }
        };
    };
}

// convert wide string to UTF-8
inline std::string narrow(const wchar_t* ws, int ns = -1)
{
    std::string s;

    int n;
    if (ns == -1) {
        n = WideCharToMultiByte(CP_UTF8, 0, ws, ns, 0, 0, 0, 0);
    }
    else {
        n = ns;
    }
    s.reserve(n);
    
    int rc = WideCharToMultiByte(CP_UTF8, 0, ws, ns, &s[0], n, 0, 0);
    ensure (rc != 0);

    return s;
}

// SQL CREATE TABLE command from a range.
/*
inline std::wstring sqlite_create_table(const xll::OPER& name, const xll::OPER& o)
{
    ensure (o.rows() >= 2);
    std::wstring s = L"CREATE TABLE ";
    s.append(name.val.str + 1, name.val.str[0]);
    s.append(L"(\n");
    std::wstring comma = L"";
    for (int i = 0; i < o.columns(); ++i) {
        const xll::OPER& oi = o(0,i);
        ensure (oi.isStr());
        s.append(comma);
        s.append(oi.val.str + 1, oi.val.str[0]);

        switch (o(1,i).xltype) {
        case xltypeNum:
            s.append(L" REAL");
            break;
        case xltypeStr:
            s.append(L" TEXT");
            break;
        case xltypeInt:
            s.append(L" INTEGER");
            break;
        default:
            ensure(!"invalid sqlite type");
        }

        comma = L", ";
    }
    s.append(L"\n);");

    return s;
}
*/

// Works like sqlite3_exec but returns an OPER.
inline xll::OPER sqlite_range(sqlite::db& db, const char* sql, bool header = false)
{
    xll::OPER o;

    sqlite::db::stmt stmt(db);
    int rc = stmt.prepare(sql);
    if (SQLITE_OK != rc)
        throw std::runtime_error(stmt.errmsg());
    
    rc = sqlite3_step(stmt);
    if (rc != SQLITE_ROW && rc != SQLITE_DONE)
        throw std::runtime_error(stmt.errmsg());
   
    int n = sqlite3_column_count(stmt);
    if (header) {
        xll::OPER head(1, n);
        for (int i = 0; i < n; ++i) {
            head[i] = xll::OPER((const XCHAR*)sqlite3_column_name16(stmt, i));
        }
        o.push_back(head);
    }
    while (SQLITE_ROW == rc) {
        xll::OPER row(1, n);
        for (int i = 0; i < n; ++i) {
            switch (sqlite3_column_type(stmt, i)) {
            case SQLITE_FLOAT:
                row[i] = sqlite3_column_double(stmt, i);
                break;
            case SQLITE_TEXT:
                row[i] = xll::OPER((const XCHAR*)sqlite3_column_text16(stmt, i));
                break;
            case SQLITE_INTEGER:
                row[i] = sqlite3_column_int(stmt, i);
                break;
            case SQLITE_NULL:
                row[i] = xll::OPER(xlerr::NA); // #NULL ???
                break;
            default:
                row[i] = xll::OPER(xlerr::NA);
            }
        }
        o.push_back(row);
        rc = sqlite3_step(stmt);
    };

    return o;
}

