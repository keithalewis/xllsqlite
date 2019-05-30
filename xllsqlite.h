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
            sqlite3_stmt** operator&()
            {
                return &pstmt;
            }
            int reset()
            {
                return sqlite3_reset(pstmt);
            }
            int step()
            {
                return sqlite3_step(pstmt);
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
            int bind(int col, const char* t, int n = -1, void(*dealloc)(void*) = SQLITE_STATIC)
            {
                return sqlite3_bind_text(pstmt, col, t, n, dealloc);
            }
            int prepare(const char* sql, int nsql = -1)
            {
                return sqlite3_prepare_v2(db, sql, nsql, &pstmt, &tail_);
            }
        };
    };
}

// convert to range
// stmt should be prepared prior to calling
inline OPER to_range(const sqlite::db::stmt& stmt)
{
    OPER o;

    // while SQLITE_ROW == stmt.next() ...

    return o;
}

