/*
** sqlite3_vb6.c
**
** VB6-friendly wrapper shim around SQLite, with full Unicode support.
**
** ============================================================
** Marshalling design
** ============================================================
**
** Inputs (VB6 -> C):
**   VB6 calls our exports with "ByVal x As Long" passing StrPtr(s).
**   StrPtr returns the address of VB6's UTF-16 buffer (a BSTR — null-
**   terminated UTF-16 with a length prefix at offset -4). We accept
**   that as `const wchar_t*` and convert UTF-16 -> UTF-8 with
**   WideCharToMultiByte before handing to SQLite.
**
**   The .bas hides this: public wrappers accept normal As String
**   parameters and call StrPtr() at the boundary.
**
** Outputs (C -> VB6):
**   String returns are wrapped in a VARIANT (VT_BSTR) and declared as
**   "As Variant" on the VB6 side. We can't use bare "As String" return
**   because VB6's Declare treats that as LPSTR (ANSI char*) — it reads
**   bytes until a NUL, which truncates UTF-16 at the first 0x00 high
**   byte. VARIANT carries a real BSTR with proper length prefix and is
**   the standard COM Automation return type, so VB6 unwraps it
**   correctly with full Unicode preserved.
**
**   For ByRef String OUT-parameters (e.g. exec_simple's pzErrMsg),
**   plain BSTR* works correctly because VB6's ByRef String marshalling
**   is BSTR-aware (unlike the return-value path).
**
** Calling convention:
**   All exports are __stdcall with undecorated names via the .def file.
**   SQLite's core compiles with its default cdecl — we don't touch
**   SQLite's calling convention. Only the thin vb6_* layer is stdcall.
**
** Handles:
**   sqlite3*, sqlite3_stmt* are 32-bit pointers (VB6 is x86), held as
**   Long in VB6.
**
** Build:
**   Compile alongside sqlite3.c (the amalgamation). Link with
**   sqlite3_vb6.def for clean undecorated stdcall exports.
*/

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <oleauto.h>     /* SysAllocStringLen, SysFreeString */
#include <stddef.h>
#include <string.h>
#include "sqlite3.h"

/* ------------------------------------------------------------------ */
/* String conversion helpers                                          */
/* ------------------------------------------------------------------ */

/*
** Convert a null-terminated UTF-16 string (from VB6's StrPtr) into a
** freshly allocated UTF-8 buffer. Caller frees with sqlite3_free.
** Returns NULL if input is NULL or allocation fails.
** Returns a valid empty string for an empty input.
*/
static char *wide_to_utf8(const wchar_t *w) {
    int wlen, ulen;
    char *out;

    if (w == NULL) return NULL;

    wlen = (int)wcslen(w);
    if (wlen == 0) {
        out = (char*)sqlite3_malloc(1);
        if (out) out[0] = '\0';
        return out;
    }

    ulen = WideCharToMultiByte(CP_UTF8, 0, w, wlen, NULL, 0, NULL, NULL);
    if (ulen <= 0) return NULL;

    out = (char*)sqlite3_malloc(ulen + 1);
    if (!out) return NULL;

    WideCharToMultiByte(CP_UTF8, 0, w, wlen, out, ulen, NULL, NULL);
    out[ulen] = '\0';
    return out;
}

/*
** Convert a UTF-8 C string into a freshly allocated BSTR. Caller owns
** the BSTR and must free with SysFreeString — but in our case it gets
** wrapped in a VARIANT and ownership transfers to VB6.
** Returns NULL if input is NULL.
*/
static BSTR utf8_to_bstr(const char *utf8) {
    int ulen, wlen;
    BSTR out;

    if (utf8 == NULL) return NULL;

    ulen = (int)strlen(utf8);
    if (ulen == 0) return SysAllocString(L"");

    wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, ulen, NULL, 0);
    if (wlen <= 0) return SysAllocString(L"");

    out = SysAllocStringLen(NULL, wlen);
    if (!out) return NULL;

    MultiByteToWideChar(CP_UTF8, 0, utf8, ulen, out, wlen);
    /* SysAllocStringLen already null-terminated at out[wlen]. */
    return out;
}

/*
** Wrap a UTF-8 string in a VARIANT of type VT_BSTR for return to VB6.
** VB6's "As Variant" return type understands this and unwraps to a
** String cleanly with full Unicode preserved.
**
** Why VARIANT instead of bare BSTR: VB6's Declare runtime treats
** "As String" returns as LPSTR (ANSI char*), which truncates UTF-16
** at the first 0x00 byte. "As Variant" doesn't have that baggage —
** it's the standard COM Automation return type and handles BSTR
** payload correctly.
**
** ABI note: VARIANT is 16 bytes, returned on x86 stdcall via a hidden
** first parameter (caller-provided buffer pointer). This is the same
** convention every IDispatch::Invoke-returning function uses; VB6
** generates the right call-site code for it automatically when the
** Declare return type is "As Variant".
**
** Returns VT_NULL variant if utf8 is NULL — VB6 sees this as Null,
** distinguishable from "" via IsNull().
*/
static VARIANT utf8_to_variant(const char *utf8) {
    VARIANT v;
    VariantInit(&v);
    if (utf8 == NULL) {
        v.vt = VT_NULL;
    } else {
        v.vt = VT_BSTR;
        v.bstrVal = utf8_to_bstr(utf8);
        if (v.bstrVal == NULL) {
            v.vt = VT_NULL;  /* allocation failure -> Null */
        }
    }
    return v;
}

/* ------------------------------------------------------------------ */
/* Library lifecycle / version                                        */
/* ------------------------------------------------------------------ */

__declspec(dllexport) VARIANT __stdcall vb6_sqlite3_libversion(void) {
    return utf8_to_variant(sqlite3_libversion());
}

__declspec(dllexport) int __stdcall vb6_sqlite3_libversion_number(void) {
    return sqlite3_libversion_number();
}

__declspec(dllexport) int __stdcall vb6_sqlite3_threadsafe(void) {
    return sqlite3_threadsafe();
}

__declspec(dllexport) int __stdcall vb6_sqlite3_initialize(void) {
    return sqlite3_initialize();
}

__declspec(dllexport) int __stdcall vb6_sqlite3_shutdown(void) {
    return sqlite3_shutdown();
}

/* ------------------------------------------------------------------ */
/* Connection: open / close                                           */
/* ------------------------------------------------------------------ */

__declspec(dllexport) int __stdcall vb6_sqlite3_open(
    const wchar_t *zFilename,
    sqlite3 **ppDb
) {
    char *utf8;
    int rc;

    if (ppDb == NULL) return SQLITE_MISUSE;
    *ppDb = NULL;

    utf8 = wide_to_utf8(zFilename);
    if (zFilename != NULL && utf8 == NULL) return SQLITE_NOMEM;

    rc = sqlite3_open(utf8 ? utf8 : "", ppDb);
    sqlite3_free(utf8);
    return rc;
}

__declspec(dllexport) int __stdcall vb6_sqlite3_open_v2(
    const wchar_t *zFilename,
    sqlite3 **ppDb,
    int flags,
    const wchar_t *zVfs
) {
    char *fn = NULL, *vfs = NULL;
    int rc;

    if (ppDb == NULL) return SQLITE_MISUSE;
    *ppDb = NULL;

    fn = wide_to_utf8(zFilename);
    if (zFilename != NULL && fn == NULL) return SQLITE_NOMEM;

    if (zVfs != NULL && zVfs[0] != L'\0') {
        vfs = wide_to_utf8(zVfs);
        if (vfs == NULL) { sqlite3_free(fn); return SQLITE_NOMEM; }
    }

    rc = sqlite3_open_v2(fn ? fn : "", ppDb, flags, vfs);
    sqlite3_free(fn);
    sqlite3_free(vfs);
    return rc;
}

__declspec(dllexport) int __stdcall vb6_sqlite3_close(sqlite3 *db) {
    if (db == NULL) return SQLITE_OK;
    return sqlite3_close(db);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_close_v2(sqlite3 *db) {
    if (db == NULL) return SQLITE_OK;
    return sqlite3_close_v2(db);
}

/* ------------------------------------------------------------------ */
/* Error reporting                                                    */
/* ------------------------------------------------------------------ */

__declspec(dllexport) int __stdcall vb6_sqlite3_errcode(sqlite3 *db) {
    if (db == NULL) return SQLITE_MISUSE;
    return sqlite3_errcode(db);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_extended_errcode(sqlite3 *db) {
    if (db == NULL) return SQLITE_MISUSE;
    return sqlite3_extended_errcode(db);
}

__declspec(dllexport) VARIANT __stdcall vb6_sqlite3_errmsg(sqlite3 *db) {
    if (db == NULL) return utf8_to_variant("invalid database handle");
    return utf8_to_variant(sqlite3_errmsg(db));
}

__declspec(dllexport) VARIANT __stdcall vb6_sqlite3_errstr(int rc) {
    return utf8_to_variant(sqlite3_errstr(rc));
}

/* ------------------------------------------------------------------ */
/* Simple exec (no callback)                                          */
/* ------------------------------------------------------------------ */

/*
** Execute SQL with no row callback. Use prepare/step for SELECTs.
** pzErrMsg receives a BSTR error message on failure, or NULL on success.
** VB6 declares this ByRef ... As String and gets a String back.
*/
__declspec(dllexport) int __stdcall vb6_sqlite3_exec_simple(
    sqlite3 *db,
    const wchar_t *zSql,
    BSTR *pzErrMsg
) {
    char *sql, *err = NULL;
    int rc;

    if (db == NULL) return SQLITE_MISUSE;
    if (pzErrMsg) *pzErrMsg = NULL;

    sql = wide_to_utf8(zSql);
    if (zSql != NULL && sql == NULL) return SQLITE_NOMEM;

    rc = sqlite3_exec(db, sql ? sql : "", NULL, NULL, &err);
    sqlite3_free(sql);

    if (err) {
        if (pzErrMsg) *pzErrMsg = utf8_to_bstr(err);
        sqlite3_free(err);
    }
    return rc;
}

/* ------------------------------------------------------------------ */
/* Prepared statements                                                */
/* ------------------------------------------------------------------ */

__declspec(dllexport) int __stdcall vb6_sqlite3_prepare_v2(
    sqlite3 *db,
    const wchar_t *zSql,
    sqlite3_stmt **ppStmt
) {
    char *sql;
    int rc;

    if (db == NULL || ppStmt == NULL) return SQLITE_MISUSE;
    *ppStmt = NULL;

    sql = wide_to_utf8(zSql);
    if (zSql != NULL && sql == NULL) return SQLITE_NOMEM;

    /* Multi-statement support is not exposed here. Add a _ex variant
    ** later if you need pzTail. */
    rc = sqlite3_prepare_v2(db, sql ? sql : "", -1, ppStmt, NULL);
    sqlite3_free(sql);
    return rc;
}

__declspec(dllexport) int __stdcall vb6_sqlite3_step(sqlite3_stmt *stmt) {
    if (stmt == NULL) return SQLITE_MISUSE;
    return sqlite3_step(stmt);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_reset(sqlite3_stmt *stmt) {
    if (stmt == NULL) return SQLITE_MISUSE;
    return sqlite3_reset(stmt);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_finalize(sqlite3_stmt *stmt) {
    if (stmt == NULL) return SQLITE_OK;
    return sqlite3_finalize(stmt);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_clear_bindings(sqlite3_stmt *stmt) {
    if (stmt == NULL) return SQLITE_MISUSE;
    return sqlite3_clear_bindings(stmt);
}

/* ------------------------------------------------------------------ */
/* Bind parameters                                                    */
/* ------------------------------------------------------------------ */

__declspec(dllexport) int __stdcall vb6_sqlite3_bind_parameter_count(sqlite3_stmt *stmt) {
    if (stmt == NULL) return 0;
    return sqlite3_bind_parameter_count(stmt);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_bind_parameter_index(
    sqlite3_stmt *stmt,
    const wchar_t *zName
) {
    char *name;
    int idx;

    if (stmt == NULL) return 0;
    name = wide_to_utf8(zName);
    if (name == NULL) return 0;

    idx = sqlite3_bind_parameter_index(stmt, name);
    sqlite3_free(name);
    return idx;
}

__declspec(dllexport) int __stdcall vb6_sqlite3_bind_null(
    sqlite3_stmt *stmt, int index
) {
    if (stmt == NULL) return SQLITE_MISUSE;
    return sqlite3_bind_null(stmt, index);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_bind_int(
    sqlite3_stmt *stmt, int index, int value
) {
    if (stmt == NULL) return SQLITE_MISUSE;
    return sqlite3_bind_int(stmt, index, value);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_bind_int64(
    sqlite3_stmt *stmt, int index, sqlite3_int64 value
) {
    if (stmt == NULL) return SQLITE_MISUSE;
    return sqlite3_bind_int64(stmt, index, value);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_bind_double(
    sqlite3_stmt *stmt, int index, double value
) {
    if (stmt == NULL) return SQLITE_MISUSE;
    return sqlite3_bind_double(stmt, index, value);
}

/*
** SQLITE_TRANSIENT tells SQLite to copy the bytes immediately, so we
** can free our UTF-8 buffer on return.
*/
__declspec(dllexport) int __stdcall vb6_sqlite3_bind_text(
    sqlite3_stmt *stmt, int index, const wchar_t *value
) {
    char *utf8;
    int rc;

    if (stmt == NULL) return SQLITE_MISUSE;
    if (value == NULL) return sqlite3_bind_null(stmt, index);

    utf8 = wide_to_utf8(value);
    if (utf8 == NULL) return SQLITE_NOMEM;

    rc = sqlite3_bind_text(stmt, index, utf8, -1, SQLITE_TRANSIENT);
    sqlite3_free(utf8);
    return rc;
}

/*
** Bind a binary blob. VB6 passes VarPtr(arr(LBound)) as pData and
** the byte count as length.
*/
__declspec(dllexport) int __stdcall vb6_sqlite3_bind_blob(
    sqlite3_stmt *stmt, int index, const void *pData, int length
) {
    if (stmt == NULL) return SQLITE_MISUSE;
    if (pData == NULL || length < 0) {
        return sqlite3_bind_null(stmt, index);
    }
    return sqlite3_bind_blob(stmt, index, pData, length, SQLITE_TRANSIENT);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_bind_zeroblob(
    sqlite3_stmt *stmt, int index, int n
) {
    if (stmt == NULL) return SQLITE_MISUSE;
    return sqlite3_bind_zeroblob(stmt, index, n);
}

/* ------------------------------------------------------------------ */
/* Column metadata                                                    */
/* ------------------------------------------------------------------ */

__declspec(dllexport) int __stdcall vb6_sqlite3_column_count(sqlite3_stmt *stmt) {
    if (stmt == NULL) return 0;
    return sqlite3_column_count(stmt);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_data_count(sqlite3_stmt *stmt) {
    if (stmt == NULL) return 0;
    return sqlite3_data_count(stmt);
}

__declspec(dllexport) VARIANT __stdcall vb6_sqlite3_column_name(
    sqlite3_stmt *stmt, int iCol
) {
    if (stmt == NULL) return utf8_to_variant(NULL);
    return utf8_to_variant(sqlite3_column_name(stmt, iCol));
}

__declspec(dllexport) VARIANT __stdcall vb6_sqlite3_column_decltype(
    sqlite3_stmt *stmt, int iCol
) {
    if (stmt == NULL) return utf8_to_variant(NULL);
    return utf8_to_variant(sqlite3_column_decltype(stmt, iCol));
}

__declspec(dllexport) int __stdcall vb6_sqlite3_column_type(
    sqlite3_stmt *stmt, int iCol
) {
    if (stmt == NULL) return SQLITE_NULL;
    return sqlite3_column_type(stmt, iCol);
}

/* ------------------------------------------------------------------ */
/* Column value extraction                                            */
/* ------------------------------------------------------------------ */

__declspec(dllexport) int __stdcall vb6_sqlite3_column_int(
    sqlite3_stmt *stmt, int iCol
) {
    if (stmt == NULL) return 0;
    return sqlite3_column_int(stmt, iCol);
}

__declspec(dllexport) sqlite3_int64 __stdcall vb6_sqlite3_column_int64(
    sqlite3_stmt *stmt, int iCol
) {
    if (stmt == NULL) return 0;
    return sqlite3_column_int64(stmt, iCol);
}

__declspec(dllexport) double __stdcall vb6_sqlite3_column_double(
    sqlite3_stmt *stmt, int iCol
) {
    if (stmt == NULL) return 0.0;
    return sqlite3_column_double(stmt, iCol);
}

__declspec(dllexport) VARIANT __stdcall vb6_sqlite3_column_text(
    sqlite3_stmt *stmt, int iCol
) {
    const unsigned char *txt;
    if (stmt == NULL) return utf8_to_variant(NULL);
    txt = sqlite3_column_text(stmt, iCol);
    if (txt == NULL) return utf8_to_variant(NULL);
    return utf8_to_variant((const char*)txt);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_column_bytes(
    sqlite3_stmt *stmt, int iCol
) {
    if (stmt == NULL) return 0;
    return sqlite3_column_bytes(stmt, iCol);
}

/*
** Two-step blob fetch from VB6:
**   1) n = sqlite3_column_bytes(stmt, iCol)
**   2) ReDim arr(0 To n-1) As Byte
**   3) sqlite3_column_blob_copy(stmt, iCol, VarPtr(arr(0)), n)
** Returns bytes copied, or -1 on bad arguments, 0 on NULL/empty.
*/
__declspec(dllexport) int __stdcall vb6_sqlite3_column_blob_copy(
    sqlite3_stmt *stmt, int iCol, void *pDest, int destSize
) {
    const void *src;
    int srcLen;

    if (stmt == NULL || pDest == NULL || destSize <= 0) return -1;

    src = sqlite3_column_blob(stmt, iCol);
    srcLen = sqlite3_column_bytes(stmt, iCol);
    if (src == NULL || srcLen <= 0) return 0;

    if (srcLen > destSize) srcLen = destSize;
    memcpy(pDest, src, (size_t)srcLen);
    return srcLen;
}

/* ------------------------------------------------------------------ */
/* Misc useful bits                                                   */
/* ------------------------------------------------------------------ */

__declspec(dllexport) sqlite3_int64 __stdcall vb6_sqlite3_last_insert_rowid(sqlite3 *db) {
    if (db == NULL) return 0;
    return sqlite3_last_insert_rowid(db);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_changes(sqlite3 *db) {
    if (db == NULL) return 0;
    return sqlite3_changes(db);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_total_changes(sqlite3 *db) {
    if (db == NULL) return 0;
    return sqlite3_total_changes(db);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_busy_timeout(sqlite3 *db, int ms) {
    if (db == NULL) return SQLITE_MISUSE;
    return sqlite3_busy_timeout(db, ms);
}

__declspec(dllexport) int __stdcall vb6_sqlite3_get_autocommit(sqlite3 *db) {
    if (db == NULL) return 0;
    return sqlite3_get_autocommit(db);
}

__declspec(dllexport) void __stdcall vb6_sqlite3_interrupt(sqlite3 *db) {
    if (db != NULL) sqlite3_interrupt(db);
}

/* End of sqlite3_vb6.c */
