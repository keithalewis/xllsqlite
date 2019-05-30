// xllsqlite.h - sqlite3 wrapper
#pragma once
#include "sqlite-amalgamation-3280000/sqlite3.h"
#include "xll12/xll/xll.h"

namespace sqlite {

    class db {
        sqlite3* pdb;
        char* errmsg_;
    public:
        db(const wchar_t* file)
            : errmsg_(0)
        {
            ensure(SQLITE_OK == sqlite3_open16(file, &pdb));
        }
        db(const db&) = delete;
        db& operator=(const db&) = delete;
        ~db()
        {
            sqlite3_close(pdb);
        }
        operator sqlite3*() {
            return pdb;
        }
        const char* errmsg() const
        {
            return errmsg_;
        }

        int exec(const char* sql, int(*cb)(void*, int, char**, char**) = 0, void* arg = 0)
        {
            if (errmsg_)
                sqlite3_free(errmsg_);

            return sqlite3_exec(pdb, sql, cb, arg, &errmsg_);
        }
        class stmt {
            sqlite::db& db;
            sqlite3_stmt* pstmt;
            const wchar_t* tail_;
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
            operator sqlite3_stmt*()
            {
                return pstmt;
            }
            int reset()
            {
                return sqlite3_reset(pstmt);
            }
            int step()
            {
                return sqlite3_step(pstmt);
            }
            const wchar_t* tail() const
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
            int bind(int col, const wchar_t* t, int n = -1, void(*dealloc)(void*) = SQLITE_STATIC)
            {
                return sqlite3_bind_text16(pstmt, col, t, n, dealloc);
            }
            int prepare(const wchar_t* sql, int nsql = -1)
            {
                return sqlite3_prepare16_v2(db, sql, nsql, &pstmt, (const void**)&tail_);
            }
        };
    };
}

// SQL CREATE TABLE command from a range.
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

// Works like sqlite3_exec but returns an OPER.
inline xll::OPER sqlite_range(sqlite::db& db, const XCHAR* sql, bool header = false)
{
    xll::OPER o;

    sqlite::db::stmt stmt(db);
    ensure (SQLITE_OK == stmt.prepare(sql));
    
    if (SQLITE_ROW != stmt.step())
        return o;
   
    int n = sqlite3_column_count(stmt);
    if (header) {
        xll::OPER head(1, n);
        for (int i = 0; i < n; ++i) {
            head[i] = xll::OPER((const XCHAR*)sqlite3_column_name16(stmt, i));
        }
        o.push_back(head);
    }
    do {
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
    } while (SQLITE_ROW == stmt.step());

    return o;
}

