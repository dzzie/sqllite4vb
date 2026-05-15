DROP TABLE IF EXISTS "users";
CREATE TABLE "users" ("id" INTEGER, "name" TEXT, "age" INTEGER, "score" FLOAT, "notes" TEXT, "joined" TEXT);
INSERT INTO "users" ("id", "name", "age", "score", "notes", "joined") VALUES (1, 'alice', 30, 99.5, 'a normal row', '2024-01-15');
INSERT INTO "users" ("id", "name", "age", "score", "notes", "joined") VALUES (2, 'O''Brien', 45, 87, 'apostrophes work', '2024-02-03');
INSERT INTO "users" ("id", "name", "age", "score", "notes", "joined") VALUES (3, 'Smith, John', 27, -3.14, 'embedded "quotes" and a comma', '2024-03-22');
INSERT INTO "users" ("id", "name", "age", "score", "notes", "joined") VALUES (4, 'bob', 52, 1.5, 'sql-injection-shaped: ''); DROP TABLE', '2024-04-01');
INSERT INTO "users" ("id", "name", "age", "score", "notes", "joined") VALUES (5, NULL, 19, NULL, 'nullable name and score', '2024-05-10');
INSERT INTO "users" ("id", "name", "age", "score", "notes", "joined") VALUES (6, 'carol', 38, 75, 'clean integer score', '2024-06-14');
INSERT INTO "users" ("id", "name", "age", "score", "notes", "joined") VALUES (7, 'line one
line two', 41, 88.8, 'quoted field with newline', '2024-07-04');
INSERT INTO "users" ("id", "name", "age", "score", "notes", "joined") VALUES (8, 'dave', 64, NULL, NULL, '2024-08-22');
