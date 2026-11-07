// MatOffice Sheets – dynamic array functions: results spill, an empty result is #CALC!, missing cells are #N/A.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluate, workbook } from './harness.mjs';

const check = (cases, cells = {}) => { for (const [f, want] of cases) assert.deepEqual(evaluate(f, cells), want, f); };

const DATA = { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30, C1: 'apple', C2: 'Banana', C3: 'cherry' };

test('UNIQUE: rows (or columns) without case, exactly_once', () => {
  check([
    ['UNIQUE({1;2;2;3})', [[1], [2], [3]]],
    ['UNIQUE({"a";"A";"b"})', [['a'], ['b']]],
    ['UNIQUE({1;1;2},,TRUE)', 2],
    ['UNIQUE({1,1,2},TRUE)', [[1, 2]]],
    ['UNIQUE({1;1},,TRUE)', '#CALC!'],
    ['UNIQUE({1,"a";1,"A";2,"b"})', [[1, 'a'], [2, 'b']]],
    ['UNIQUE(A1:A3*0)', 0],
  ], DATA);
});

test('SORT and SORTBY: Excel order, several keys, columns', () => {
  check([
    ['SORT({3;1;2})', [[1], [2], [3]]],
    ['SORT({3;1;2},1,-1)', [[3], [2], [1]]],
    ['SORT({"b";"A";"c"})', [['A'], ['b'], ['c']]],
    ['SORT({"x";2;TRUE;1})', [[1], [2], ['x'], [true]]],
    ['SORT({3,1;1,2;2,0},2)', [[2, 0], [3, 1], [1, 2]]],
    ['SORT({"b",2;"a",1;"b",1},{1,2})', [['a', 1], ['b', 1], ['b', 2]]],
    ['SORT({"b",2;"a",1;"b",1},{1,2},{1,-1})', [['a', 1], ['b', 2], ['b', 1]]],
    ['SORT({3,1,2},1,1,TRUE)', [[1, 2, 3]]],
    ['SORT({1;2},3)', '#VALUE!'],
    ['SORT({1;2},1,2)', '#VALUE!'],
    ['SORTBY({"a";"b";"c"},{3;1;2})', [['b'], ['c'], ['a']]],
    ['SORTBY({"a";"b";"c"},{3;1;2},-1)', [['a'], ['c'], ['b']]],
    ['SORTBY({"a","b"},{2,1})', [['b', 'a']]],
    ['SORTBY(C1:C3,B1:B3,-1)', [['cherry'], ['Banana'], ['apple']]],
    ['SORTBY({"a";"b";"c"},{1;1;0},1,{3;2;1},1)', [['c'], ['b'], ['a']]],
    ['SORTBY({"a";"b"},{1;2;3})', '#VALUE!'],
  ], DATA);
});

test('FILTER keeps rows or columns and has an if_empty', () => {
  check([
    ['FILTER(A1:A3,A1:A3>1)', [[2], [3]]],
    ['FILTER(A1:C3,B1:B3>10)', [[2, 20, 'Banana'], [3, 30, 'cherry']]],
    ['FILTER({1,2,3},{TRUE,FALSE,TRUE})', [[1, 3]]],
    ['FILTER(C1:C3,(A1:A3>1)*(B1:B3<30))', 'Banana'],
    ['FILTER(A1:A3,A1:A3>5)', '#CALC!'],
    ['FILTER(A1:A3,A1:A3>5,"none")', 'none'],
    ['FILTER(A1:A3,{1;0})', '#VALUE!'],
    ['FILTER({1;2},{"a";"b"})', '#VALUE!'],
    ['FILTER({1;2},{1;#N/A})', '#N/A'],
  ], DATA);
});

test('SEQUENCE and RANDARRAY make new arrays', () => {
  check([
    ['SEQUENCE(3)', [[1], [2], [3]]],
    ['SEQUENCE(2,3)', [[1, 2, 3], [4, 5, 6]]],
    ['SEQUENCE(2,2,10,-1)', [[10, 9], [8, 7]]],
    ['SEQUENCE(3,,10)', [[10], [11], [12]]],
    ['SEQUENCE(1,1,5)', 5],
    ['SUM(SEQUENCE(100))', 5050],
    ['SEQUENCE(0)', '#CALC!'],
    ['SEQUENCE(-1)', '#VALUE!'],
    ['RANDARRAY(1,1,5,1)', '#VALUE!'],
  ]);
  const r = evaluate('RANDARRAY(2,3)');
  assert.equal(r.length, 2);
  assert.ok(r.flat().every(x => x >= 0 && x < 1) && r[0].length === 3);
  const dice = evaluate('RANDARRAY(20,1,1,6,TRUE)');
  assert.ok(dice.flat().every(x => Number.isInteger(x) && x >= 1 && x <= 6));
  const wb = workbook({ Sheet1: { A1: '=SEQUENCE(3)', B1: '=SUM(A1:A3)', C1: '=A2*10' } });
  assert.equal(wb.get('A3'), 3);
  assert.equal(wb.get('B1'), 6);
  assert.equal(wb.get('C1'), 20);
});

test('VSTACK, HSTACK, TAKE, DROP, CHOOSECOLS and CHOOSEROWS', () => {
  check([
    ['VSTACK({1,2},{3,4})', [[1, 2], [3, 4]]],
    ['VSTACK({1,2},{3})', [[1, 2], [3, '#N/A']]],
    ['VSTACK(A1:A2,B1)', [[1], [2], [10]]],
    ['HSTACK({1;2},{3;4})', [[1, 3], [2, 4]]],
    ['HSTACK({1;2},3)', [[1, 3], [2, '#N/A']]],
    ['TAKE(SEQUENCE(5),2)', [[1], [2]]],
    ['TAKE(SEQUENCE(5),-2)', [[4], [5]]],
    ['TAKE({1,2,3;4,5,6},,2)', [[1, 2], [4, 5]]],
    ['TAKE({1,2;3,4},1,-1)', 2],
    ['TAKE({1,2},0)', '#CALC!'],
    ['DROP(SEQUENCE(5),2)', [[3], [4], [5]]],
    ['DROP(SEQUENCE(5),-3)', [[1], [2]]],
    ['DROP({1,2;3,4},1,1)', 4],
    ['DROP({1,2},1)', '#CALC!'],
    ['CHOOSECOLS({1,2,3;4,5,6},1,3)', [[1, 3], [4, 6]]],
    ['CHOOSECOLS({1,2,3},-1)', 3],
    ['CHOOSECOLS({1,2},3)', '#VALUE!'],
    ['CHOOSEROWS({1;2;3},{3,1})', [[3], [1]]],
    ['CHOOSEROWS({1;2},0)', '#VALUE!'],
  ], DATA);
});

test('TOCOL, TOROW, TRANSPOSE, WRAPROWS, WRAPCOLS and EXPAND', () => {
  check([
    ['TOCOL({1,2;3,4})', [[1], [2], [3], [4]]],
    ['TOCOL({1,2;3,4},0,TRUE)', [[1], [3], [2], [4]]],
    ['TOCOL({1,#N/A;3,4},2)', [[1], [3], [4]]],
    ['TOCOL(A1:A5,1)', [[1], [2], [3]]],
    ['TOROW({1;2})', [[1, 2]]],
    ['TOROW({1;2},4)', '#VALUE!'],
    ['TRANSPOSE({1,2,3})', [[1], [2], [3]]],
    ['TRANSPOSE(A1:B2)', [[1, 2], [10, 20]]],
    ['WRAPROWS(SEQUENCE(5),2)', [[1, 2], [3, 4], [5, '#N/A']]],
    ['WRAPROWS({1,2,3},2,0)', [[1, 2], [3, 0]]],
    ['WRAPCOLS({1,2,3,4,5},2)', [[1, 3, 5], [2, 4, '#N/A']]],
    ['WRAPROWS({1,2;3,4},2)', '#VALUE!'],
    ['WRAPROWS({1},0)', '#NUM!'],
    ['EXPAND({1,2},2,3,0)', [[1, 2, 0], [0, 0, 0]]],
    ['EXPAND({1},2)', [[1], ['#N/A']]],
    ['EXPAND({1,2},1,1)', '#VALUE!'],
  ], DATA);
});
