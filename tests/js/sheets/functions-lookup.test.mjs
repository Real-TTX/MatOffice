// MatOffice Sheets – lookup & reference functions: exact and approximate matches, wildcards, references as results, spills.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluate, workbook } from './harness.mjs';

const check = (cases, cells = {}) => { for (const [f, want] of cases) assert.deepEqual(evaluate(f, cells), want, f); };

const TABLE = {
  A1: 'Name', B1: 'Qty', C1: 'Price',
  A2: 'apple', B2: 3, C2: 1.5,
  A3: 'Banana', B3: 5, C3: 0.25,
  A4: 'cherry', B4: 7, C4: 4,
  E1: 10, E2: 20, E3: 30, E4: 40, E5: 50,
  F1: 'a', F2: 'b', F3: 'c', F4: 'd', F5: 'e',
  G1: '=SUM(B2:B4)',
};

test('VLOOKUP and HLOOKUP: exact (no case, wildcards) and approximate matches', () => {
  check([
    ['VLOOKUP("banana",A2:C4,2,FALSE)', 5],
    ['VLOOKUP("ch*",A2:C4,3,FALSE)', 4],
    ['VLOOKUP("zz",A2:C4,2,FALSE)', '#N/A'],
    ['VLOOKUP("apple",A2:C4,4,FALSE)', '#REF!'],
    ['VLOOKUP("apple",A2:C4,0,FALSE)', '#VALUE!'],
    ['VLOOKUP(35,E1:F5,2)', 'c'],
    ['VLOOKUP(35,E1:F5,2,TRUE)', 'c'],
    ['VLOOKUP(30,E1:F5,2,0)', 'c'],
    ['VLOOKUP(5,E1:F5,2)', '#N/A'],
    ['VLOOKUP(99,E1:F5,2)', 'e'],
    ['VLOOKUP("3",E1:F5,2,FALSE)', '#N/A'],
    ['VLOOKUP(E1:E2,E1:F5,2,FALSE)', [['a'], ['b']]],
    ['VLOOKUP("apple",A2:C4,{2,3},FALSE)', [[3, 1.5]]],
    ['VLOOKUP("Banana",A:C,3,FALSE)', 0.25],
    ['HLOOKUP("Qty",A1:C4,2,FALSE)', 3],
    ['HLOOKUP("Price",A1:C4,4,FALSE)', 4],
    ['HLOOKUP("P*",A1:C4,3,FALSE)', 0.25],
    ['HLOOKUP("x",A1:C4,2,FALSE)', '#N/A'],
    ['HLOOKUP("Qty",A1:C4,5,FALSE)', '#REF!'],
    ['HLOOKUP(25,{10,20,30;"a","b","c"},2)', 'b'],
  ], TABLE);
});

test('XLOOKUP: match and search modes, not found, several columns and references', () => {
  check([
    ['XLOOKUP("cherry",A2:A4,C2:C4)', 4],
    ['XLOOKUP("CHERRY",A2:A4,B2:C4)', [[7, 4]]],
    ['XLOOKUP("zz",A2:A4,C2:C4)', '#N/A'],
    ['XLOOKUP("zz",A2:A4,C2:C4,"none")', 'none'],
    ['XLOOKUP(35,E1:E5,F1:F5)', '#N/A'],
    ['XLOOKUP(35,E1:E5,F1:F5,,-1)', 'c'],
    ['XLOOKUP(35,E1:E5,F1:F5,,1)', 'd'],
    ['XLOOKUP("b?nana",A2:A4,B2:B4,,2)', 5],
    ['XLOOKUP("b?nana",A2:A4,B2:B4)', '#N/A'],
    ['XLOOKUP(30,E1:E5,F1:F5,,0,2)', 'c'],
    ['XLOOKUP(35,E1:E5,F1:F5,,1,2)', 'd'],
    ['XLOOKUP(35,E1:E5,F1:F5,,-1,2)', 'c'],
    ['XLOOKUP(35,{50,40,30,20,10},{"e","d","c","b","a"},,1,-2)', 'd'],
    ['XLOOKUP(35,{50,40,30,20,10},{"e","d","c","b","a"},,-1,-2)', 'c'],
    ['XLOOKUP(1,{1,2,1},{"a","b","c"},,0,-1)', 'c'],
    ['XLOOKUP("Price",A1:C1,A4:C4)', 4],
    ['XLOOKUP({"apple";"cherry"},A2:A4,B2:C4)', [[3, 1.5], [7, 4]]],
    ['SUM(XLOOKUP("apple",A2:A4,B2:B4):XLOOKUP("cherry",A2:A4,B2:B4))', 15],
    ['XLOOKUP(1,A2:B3,C2:C3)', '#VALUE!'],
    ['XLOOKUP(10,E1:E5,F1:F4)', '#VALUE!'],
    ['XLOOKUP(10,E1:E5,F1:F5,,3)', '#VALUE!'],
  ], TABLE);
});

test('MATCH, XMATCH and LOOKUP', () => {
  check([
    ['MATCH(39,{25,38,40,41})', 2],
    ['MATCH(41,{25,38,40,41},0)', 4],
    ['MATCH(40,{25,38,40,41},-1)', '#N/A'],
    ['MATCH("b*",A2:A4,0)', 2],
    ['MATCH("BANANA",A2:A4,0)', 2],
    ['MATCH(35,E1:E5)', 3],
    ['MATCH(35,{50,40,30,20,10},-1)', 2],
    ['MATCH(2,{1,2,2,2,3},1)', 4],
    ['MATCH(2,{1,"x",2,"y",3})', 3],
    ['MATCH("1",{1,2},0)', '#N/A'],
    ['MATCH(5,A1:C2,0)', '#N/A'],
    ['MATCH(E1:E2,E1:E5,0)', [[1], [2]]],
    ['XMATCH(4,{5,4,3,2,1})', 2],
    ['XMATCH(4.5,{5,4,3,2,1},1)', 1],
    ['XMATCH("Gra*",{"Apple","Grape","Pear"},2)', 2],
    ['XMATCH("Pea?",{"Apple","Grape","Pear"},2)', 3],
    ['XMATCH(3,{1,3,3,5},0,-1)', 3],
    ['XMATCH(9,{1,2})', '#N/A'],
    ['LOOKUP(4.19,{4.14,4.19,5.17,5.77,6.39},{"red","orange","yellow","green","blue"})', 'orange'],
    ['LOOKUP(5.75,{4.14,4.19,5.17,5.77,6.39},{"red","orange","yellow","green","blue"})', 'yellow'],
    ['LOOKUP(7.66,{4.14,4.19,5.17,5.77,6.39},{"red","orange","yellow","green","blue"})', 'blue'],
    ['LOOKUP(0,{4.14,4.19,5.17,5.77,6.39},{"red","orange","yellow","green","blue"})', '#N/A'],
    ['LOOKUP("C",{"a","b","c","d";1,2,3,4})', 3],
    ['LOOKUP("bump",{"a",1;"b",2;"c",3})', 2],
    ['LOOKUP(2,1/(A2:A4<>"apple"),B2:B4)', 7],
    ['LOOKUP(9.99E+307,B:B)', 7],
    ['LOOKUP(25,E1:E5,F1)', 'b'],
  ], TABLE);
});

test('INDEX gives references (whole rows/columns with 0) and values of arrays', () => {
  check([
    ['INDEX({1,2;3,4},0,2)', [[2], [4]]],
    ['INDEX({1,2;3,4},2,2)', 4],
    ['INDEX({1,2,3},2)', 2],
    ['INDEX(A2:C4,2,3)', 0.25],
    ['INDEX(A2:C4,2)', [['Banana', 5, 0.25]]],
    ['INDEX(A2:C4,,2)', [[3], [5], [7]]],
    ['INDEX(A2:A4,3)', 'cherry'],
    ['INDEX(A1:C1,2)', 'Qty'],
    ['INDEX(A:A,3)', 'Banana'],
    ['INDEX(A2:C4,4,1)', '#REF!'],
    ['INDEX(A2:C4,-1,1)', '#VALUE!'],
    ['INDEX(A2:C4,1,1,2)', '#REF!'],
    ['SUM(INDEX(B2:C4,0,1))', 15],
    ['SUM(B2:INDEX(B2:B4,2))', 8],
    ['ROW(INDEX(A1:A9,3))', 3],
    ['INDEX(E1:E5,{1;3})', [[10], [30]]],
    ['INDEX(A2:A4,MATCH(MAX(B2:B4),B2:B4,0))', 'cherry'],
  ], TABLE);
});

test('CHOOSE is lazy and builds arrays from an array of indexes', () => {
  check([
    ['CHOOSE(2,"1st","2nd","3rd")', '2nd'],
    ['CHOOSE(4,"a","b")', '#VALUE!'],
    ['CHOOSE(0,"a")', '#VALUE!'],
    ['CHOOSE(1.9,"a","b")', 'a'],
    ['CHOOSE(2,1/0,"ok")', 'ok'],
    ['SUM(CHOOSE(2,E1:E2,E1:E5))', 150],
    ['CHOOSE({1,2},"a","b")', [['a', 'b']]],
    ['VLOOKUP(7,CHOOSE({1,2},B2:B4,A2:A4),2,FALSE)', 'cherry'],
  ], TABLE);
});

test('ROW, COLUMN, ROWS, COLUMNS and AREAS', () => {
  check([
    ['ROW()', 1000],
    ['ROW(C10)', 10],
    ['ROW(A2:A4)', [[2], [3], [4]]],
    ['SUM(ROW(A1:A3))', 6],
    ['ROW("x")', '#VALUE!'],
    ['COLUMN()', 702],
    ['COLUMN(D10)', 4],
    ['COLUMN(B1:D1)', [[2, 3, 4]]],
    ['ROWS(A1:C4)', 4],
    ['ROWS({1,2;3,4;5,6})', 3],
    ['ROWS(5)', 1],
    ['ROWS(A:A)', 1048576],
    ['COLUMNS(A1:C4)', 3],
    ['COLUMNS({1,2,3})', 3],
    ['COLUMNS(1:1)', 16384],
    ['AREAS(A1:C3)', 1],
  ], TABLE);
  const wb = workbook({ Sheet1: { A1: 'x', A5: '=ROW()-1', B1: '=IF(A1="","",ROW()-0)' } });
  assert.equal(wb.get('A5'), 4);
  assert.equal(wb.get('B1'), 1);
});

test('OFFSET and INDIRECT build references', () => {
  check([
    ['OFFSET(A1,2,1)', 5],
    ['OFFSET(A1,1,1,3,2)', [[3, 1.5], [5, 0.25], [7, 4]]],
    ['SUM(OFFSET(B2,0,0,3))', 15],
    ['SUM(OFFSET(B4,0,0,-2))', 12],
    ['ROWS(OFFSET(A1,0,0,5,2))', 5],
    ['OFFSET(A1,-1,0)', '#REF!'],
    ['OFFSET(A1,0,0,0)', '#REF!'],
    ['OFFSET("x",1,1)', '#VALUE!'],
    ['INDIRECT("B3")', 5],
    ['INDIRECT("A"&3)', 'Banana'],
    ['SUM(INDIRECT("B2:B4"))', 15],
    ['INDIRECT("R3C2",FALSE)', 5],
    ['INDIRECT("R[-997]C[-700]",FALSE)', 5],
    ['INDIRECT("nope")', '#REF!'],
    ['INDIRECT("R0C1",FALSE)', '#REF!'],
  ], TABLE);
  const wb = workbook({
    Sheet1: { A1: '=INDIRECT("\'Other sheet\'!B2")', A2: '=SUM(INDIRECT("Amounts"))', A3: '=INDIRECT("Other sheet!R2C2",FALSE)', A4: '=OFFSET(Rate,1,0)' },
    'Other sheet': { B2: 42, B3: 8 },
  }, [{ name: 'Amounts', ref: "'Other sheet'!$B$2:$B$3" }, { name: 'Rate', ref: "'Other sheet'!$B$2" }]);
  assert.equal(wb.get('A1'), 42);
  assert.equal(wb.get('A2'), 50);
  assert.equal(wb.get('A3'), 42);
  assert.equal(wb.get('A4'), 8);
  wb.set({ B2: 1 }, 'Other sheet');
  assert.equal(wb.get('A1'), 1);
});

test('ADDRESS, HYPERLINK and FORMULATEXT', () => {
  check([
    ['ADDRESS(2,3)', '$C$2'],
    ['ADDRESS(2,3,2)', 'C$2'],
    ['ADDRESS(2,3,3)', '$C2'],
    ['ADDRESS(2,3,4)', 'C2'],
    ['ADDRESS(2,3,2,FALSE)', 'R2C[3]'],
    ['ADDRESS(2,3,1,FALSE,"[Book1]Sheet1")', "'[Book1]Sheet1'!R2C3"],
    ['ADDRESS(2,3,1,FALSE,"EXCEL SHEET")', "'EXCEL SHEET'!R2C3"],
    ['ADDRESS(5,28,1,TRUE,"Sheet2")', 'Sheet2!$AB$5'],
    ['ADDRESS(1,1,,FALSE)', 'R1C1'],
    ['ADDRESS(0,1)', '#VALUE!'],
    ['ADDRESS(1,16385)', '#VALUE!'],
    ['HYPERLINK("https://example.com","Example")', 'Example'],
    ['HYPERLINK("https://example.com")', 'https://example.com'],
    ['FORMULATEXT(G1)', '=SUM(B2:B4)'],
    ['FORMULATEXT(A1)', '#N/A'],
  ], TABLE);
});
