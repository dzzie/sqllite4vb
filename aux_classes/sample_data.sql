-- Sample data for cSQLiteTable testing.
-- Covers the cases that exercise the class:
--   users        - the canonical example from the class headers
--   events       - mixed types including REAL and ISO date strings
--   blobs        - exercises the BLOB path (hex round-trip in JSON)
--   nullable     - rows with NULLs in various columns

-- ---------------------------------------------------------------
-- users: the "SELECT name, age FROM users WHERE age > 30" example
-- ---------------------------------------------------------------
DROP TABLE IF EXISTS users;
CREATE TABLE users (
    id      INTEGER PRIMARY KEY,
    name    TEXT NOT NULL,
    age     INTEGER NOT NULL
);

INSERT INTO users (name, age) VALUES
    ('alice',   30),
    ('bob',     45),
    ('carol',   27),
    ('dave',    52),
    ('eve',     38),
    ('frank',   19),
    ('grace',   64),
    ('henry',   41);

-- ---------------------------------------------------------------
-- events: REAL column (sqlFloat), TEXT timestamps, INTEGER counts
-- Useful for confirming type hints come back correctly.
-- ---------------------------------------------------------------
DROP TABLE IF EXISTS events;
CREATE TABLE events (
    id          INTEGER PRIMARY KEY,
    occurred_at TEXT    NOT NULL,    -- ISO 8601
    severity    REAL    NOT NULL,    -- 0.0 .. 10.0
    hit_count   INTEGER NOT NULL,
    source      TEXT
);

INSERT INTO events (occurred_at, severity, hit_count, source) VALUES
    ('2026-05-10T08:14:22', 2.5,  1,  'sensor-A'),
    ('2026-05-10T09:01:07', 7.25, 4,  'sensor-A'),
    ('2026-05-11T14:33:50', 9.9,  17, 'sensor-B'),
    ('2026-05-12T03:12:00', 0.1,  1,  NULL),
    ('2026-05-13T22:48:15', 5.0,  3,  'sensor-C');

-- ---------------------------------------------------------------
-- blobs: exercises the BLOB path. SQLite's X'...' literal syntax
-- inserts raw bytes. cSQLiteTable will surface these as Byte arrays
-- and serialize to JSON as {"$blob":"<hex>"}.
-- ---------------------------------------------------------------
DROP TABLE IF EXISTS blobs;
CREATE TABLE blobs (
    id      INTEGER PRIMARY KEY,
    label   TEXT NOT NULL,
    payload BLOB NOT NULL
);

INSERT INTO blobs (label, payload) VALUES
    ('hello-ascii',  X'48656C6C6F'),                        -- "Hello"
    ('png-header',   X'89504E470D0A1A0A'),                  -- PNG magic
    ('zero-fill',    X'00000000'),
    ('mixed',        X'DEADBEEFCAFEBABE');

-- ---------------------------------------------------------------
-- nullable: every column nullable. Useful for confirming the
-- "first non-NULL wins" type-hint logic in LoadFromResults, and
-- that NULLs round-trip through JSON as `null`.
-- ---------------------------------------------------------------
DROP TABLE IF EXISTS nullable;
CREATE TABLE nullable (
    id      INTEGER PRIMARY KEY,
    a_int   INTEGER,
    a_real  REAL,
    a_text  TEXT,
    a_blob  BLOB
);

INSERT INTO nullable (a_int, a_real, a_text, a_blob) VALUES
    (NULL, NULL,  NULL,      NULL),                 -- all-null row
    (42,   NULL,  'forty-two', NULL),
    (NULL, 3.14,  NULL,      X'010203'),
    (7,    2.718, 'mixed',   X'FF');
