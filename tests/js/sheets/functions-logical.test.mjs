// MatOffice Sheets – logical functions: lazy branches, arrays as tests, Excel's rules for logicals, text and blanks.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluate, workbook } from './harness.mjs';

const check = (cases, cells = {}) => { for (const [f, want] of cases) assert.deepEqual(evaluate(f, cells), want, f); };

const DATA = { A1: 1, A2: 2, A3: 3, A4: 'x', A5: true, B1: 10, B2: 20, B3: 30, C1: 'apple', C2: 'Banana', C3: 'cherry', E1: '#N/A' };

test('IF: branches, empty arguments, text tests and errors', () => {
  check([
    ['IF(TRUE,1,2)', 1],
    ['IF(FALSE,1,2)', 2],
    ['IF(FALSE,1)', false],
    ['IF(TRUE,,2)', 0],
    ['IF(FALSE,1,)', 0],
    ['IF("TRUE",1,2)', 1],
    ['IF("x",1,2)', '#VALUE!'],
    ['IF(E1,1,2)', '#N/A'],
    ['IF(A6,1,2)', 2],
    ['IF(A1,"one","none")', 'one'],
    ['IF(1/0>1,1,2)', '#DIV/0!'],
    ['IF(TRUE,1,1/0)', 1],
    ['IF(FALSE,NA(),5)', 5],
    ['IF(A1=1,"",A2)', ''],
    ['SUM(IF(TRUE,A1:A3))', 6],
  ], DATA);
});

test('IF with an array test works element by element and spills', () => {
  check([
    ['IF(A1:A3>1,"y","n")', [['n'], ['y'], ['y']]],
    ['IF(A1:A3>1,B1:B3,0)', [[0], [20], [30]]],
    ['IF({TRUE,FALSE},{1,2},{3,4})', [[1, 4]]],
    ['IF(TRUE,A1:A3)', [[1], [2], [3]]],
    ['SUM(IF(A1:A3>1,B1:B3))', 50],
    ['IF({1;0;"x"},"a","b")', [['a'], ['b'], ['#VALUE!']]],
  ], DATA);
  const wb = workbook({ Sheet1: { A1: 1, A2: 5, A3: 9, B1: '=IF(A1:A3>4,"big","small")' } });
  assert.equal(wb.get('B1'), 'small');
  assert.equal(wb.get('B2'), 'big');
  assert.equal(wb.get('B3'), 'big');
});

test('IFS, IFERROR and IFNA', () => {
  check([
    ['IFS(A1>5,"a",A1>0,"b")', 'b'],
    ['IFS(FALSE,1)', '#N/A'],
    ['IFS("x",1)', '#VALUE!'],
    ['IFS(TRUE,)', 0],
    ['IFS(FALSE,1/0,TRUE,2)', 2],
    ['IFS(A1:A3>1,"big",TRUE,"small")', [['small'], ['big'], ['big']]],
    ['IFERROR(1/0,"x")', 'x'],
    ['IFERROR(5,"x")', 5],
    ['IFERROR(A6,"x")', 0],
    ['IFERROR(E1,)', 0],
    ['IFERROR({1,#N/A,3},0)', [[1, 0, 3]]],
    ['IFERROR(A1:A3/{1;0;1},"div")', [[1], ['div'], [3]]],
    ['IFNA(E1,"none")', 'none'],
    ['IFNA(1/0,"none")', '#DIV/0!'],
    ['IFNA(VLOOKUP("zz",C1:C3,1,FALSE),"missing")', 'missing'],
    ['IFNA(7,"none")', 7],
  ], DATA);
});

test('AND, OR, XOR and NOT: logicals and numbers count, text in ranges is skipped', () => {
  check([
    ['AND(TRUE,TRUE)', true],
    ['AND(TRUE,FALSE)', false],
    ['AND(A1:A5)', true],
    ['AND(A1:A3,0)', false],
    ['AND(A1:A3>0)', true],
    ['AND(A4)', '#VALUE!'],
    ['AND("x")', '#VALUE!'],
    ['AND("TRUE",1)', true],
    ['AND(E1)', '#N/A'],
    ['OR(FALSE,0)', false],
    ['OR(A1>2,A2>1)', true],
    ['OR("TRUE")', true],
    ['OR(A4)', '#VALUE!'],
    ['OR({FALSE,TRUE})', true],
    ['XOR(TRUE,TRUE,TRUE)', true],
    ['XOR(TRUE,TRUE)', false],
    ['XOR(3>0,2<9)', false],
    ['XOR(3>12,4<6)', true],
    ['NOT(TRUE)', false],
    ['NOT(0)', true],
    ['NOT(A6)', true],
    ['NOT("x")', '#VALUE!'],
    ['NOT({0,1})', [[true, false]]],
    ['TRUE()', true],
    ['FALSE()', false],
  ], DATA);
});

test('SWITCH compares like = and lifts over an array expression', () => {
  check([
    ['SWITCH(2,1,"a",2,"b","c")', 'b'],
    ['SWITCH(9,1,"a","c")', 'c'],
    ['SWITCH(9,1,"a")', '#N/A'],
    ['SWITCH("B","a",1,"b",2)', 2],
    ['SWITCH(1/0,1,2)', '#DIV/0!'],
    ['SWITCH(A1:A3,1,"one",2,"two","many")', [['one'], ['two'], ['many']]],
    ['SWITCH(WEEKDAY(DATE(2026,1,15)),1,"Sunday",5,"Thursday","?")', 'Thursday'],
    ['SWITCH(2,1,1/0,2,"ok")', 'ok'],
  ], DATA);
});
