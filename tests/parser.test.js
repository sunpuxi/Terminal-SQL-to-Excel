const assert = require("assert");
const { parseMysqlConsole, displayLength, selectColumns } = require("../app.js");

function parsed(text) {
  const result = parseMysqlConsole(text);
  assert.strictEqual(result.ok, true, result.error);
  return result;
}

const ordinary = parsed(`
+----+----------+--------+
| id | name     | status |
+----+----------+--------+
| 1  | Alice    | active |
| 2  | Bob      | closed |
+----+----------+--------+
2 rows in set (0.01 sec)`);
assert.deepStrictEqual(ordinary.data, [["id", "name", "status"], ["1", "Alice", "active"], ["2", "Bob", "closed"]]);
assert.deepStrictEqual(selectColumns(ordinary.data, [0, 2]), [["id", "status"], ["1", "active"], ["2", "closed"]]);
assert.deepStrictEqual(selectColumns(ordinary.data, [1]), [["name"], ["Alice"], ["Bob"]]);
assert.deepStrictEqual(selectColumns(ordinary.data, []), []);

const chinese = parsed(`| id | name | status |
| 1 | 张三 | 正常 |
| 2 | 李四 | 已删除 |`);
assert.strictEqual(chinese.rows[1][1], "李四");
assert.strictEqual(displayLength("李四"), 4);

const nullValue = parsed(`| id | name | email |
| 1 | Alice | NULL |`);
assert.strictEqual(nullValue.rows[0][2], "NULL");

const empty = parsed(`| id | name | status |
| 1 |      | active |`);
assert.strictEqual(empty.rows[0][1], "");

const fullConsole = parsed(`mysql> select id,name from users;
+----+-------+
| id | name  |
+----+-------+
| 1  | Alice |
| 2  | Bob   |
+----+-------+
2 rows in set (0.01 sec)

mysql>`);
assert.strictEqual(fullConsole.rowCount, 2);

assert.strictEqual(parseMysqlConsole("hello world").ok, false);
assert.match(parseMysqlConsole("| id | name |\n| 1 | Alice | active |").error, /列数与表头不一致/);

console.log("All parser tests passed.");
