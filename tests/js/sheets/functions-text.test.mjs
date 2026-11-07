// MatOffice Sheets – text functions against Excel's results (Excel's own help examples where there are some).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluate } from './harness.mjs';

const approx = (value, tol = 1e-9) => ({ approx: value, tol });

function expect(got, want, label) {
  if (want && typeof want === 'object' && 'approx' in want) {
    assert.equal(typeof got, 'number', `${label}: ${JSON.stringify(got)} is not a number`);
    assert.ok(Math.abs(got - want.approx) <= want.tol * Math.max(1, Math.abs(want.approx)), `${label}: ${got} is not ${want.approx}`);
  } else {
    assert.deepEqual(got, want, label);
  }
}

const check = (cases, cells = {}, options = {}) => { for (const [f, want] of cases) expect(evaluate(f, cells, options), want, f); };

const DATA = { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, C1: 'apple', C2: 'Banana', C3: 'cherry', E1: '#N/A' };

test('CONCAT, CONCATENATE and TEXTJOIN', () => {
  check([
    ['CONCAT(C1:C3)', 'appleBananacherry'],
    ['CONCAT("a",1,TRUE)', 'a1TRUE'],
    ['CONCAT(A1:B2)', '110220'],
    ['CONCAT(E1)', '#N/A'],
    ['CONCAT(0.1+0.2)', '0.3'],
    ['CONCATENATE("a","b",1)', 'ab1'],
    ['CONCATENATE(C1:C2,"!")', [['apple!'], ['Banana!']]],
    ['TEXTJOIN(", ",TRUE,C1:C3,"")', 'apple, Banana, cherry'],
    ['TEXTJOIN("-",FALSE,"a","","b")', 'a--b'],
    ['TEXTJOIN({"-","+"},FALSE,1,2,3)', '1-2+3'],
    ['TEXTJOIN(",",TRUE,A6:A7)', ''],
    ['TEXTJOIN(",",FALSE,A1:A3)', '1,2,3'],
    ['TEXTJOIN(",",TRUE,E1)', '#N/A'],
  ], DATA);
  assert.equal(evaluate('CONCAT(2.5,"|")', {}, { locale: 'de' }), '2,5|');
});

test('LEFT, RIGHT, MID and LEN', () => {
  check([
    ['LEFT("Sale Price",4)', 'Sale'],
    ['LEFT("Sweden")', 'S'],
    ['LEFT("abc",)', ''],
    ['LEFT("abc",-1)', '#VALUE!'],
    ['LEFT(12345,2)', '12'],
    ['RIGHT("Sale Price",5)', 'Price'],
    ['RIGHT("Stock Number")', 'r'],
    ['RIGHT("abc",0)', ''],
    ['RIGHT("abc",10)', 'abc'],
    ['MID("Fluid Flow",1,5)', 'Fluid'],
    ['MID("Fluid Flow",7,20)', 'Flow'],
    ['MID("Fluid Flow",20,5)', ''],
    ['MID("abc",0,1)', '#VALUE!'],
    ['MID("abc",1,-1)', '#VALUE!'],
    ['LEN("Phoenix, AZ")', 11],
    ['LEN("")', 0],
    ['LEN(A6)', 0],
    ['LEN(123.5)', 5],
    ['LEN(C1:C3)', [['apple'.length], ['Banana'.length], ['cherry'.length]]],
  ], DATA);
});

test('FIND (with case) and SEARCH (without case, with wildcards)', () => {
  check([
    ['FIND("M","Miriam McGovern")', 1],
    ['FIND("m","Miriam McGovern")', 6],
    ['FIND("M","Miriam McGovern",3)', 8],
    ['FIND("z","abc")', '#VALUE!'],
    ['FIND("","abc")', 1],
    ['FIND("a","abc",0)', '#VALUE!'],
    ['SEARCH("e","Statements",6)', 7],
    ['SEARCH("margin","Profit Margin")', 8],
    ['SEARCH("an*a","Banana")', 2],
    ['SEARCH("?a","Banana")', 1],
    ['SEARCH("~*","a*b")', 2],
    ['SEARCH("B","abc")', 2],
    ['SEARCH("x","abc")', '#VALUE!'],
  ]);
});

test('SUBSTITUTE, REPLACE and REPT', () => {
  check([
    ['SUBSTITUTE("Sales Data","Sales","Cost")', 'Cost Data'],
    ['SUBSTITUTE("Quarter 1, 2008","1","2",1)', 'Quarter 2, 2008'],
    ['SUBSTITUTE("Quarter 1, 2011","1","2",3)', 'Quarter 1, 2012'],
    ['SUBSTITUTE("aaa","a","b",2)', 'aba'],
    ['SUBSTITUTE("abc","","x")', 'abc'],
    ['SUBSTITUTE("abc","b","x",0)', '#VALUE!'],
    ['SUBSTITUTE("ABC","b","x")', 'ABC'],
    ['REPLACE("abcdefghijk",6,5,"*")', 'abcde*k'],
    ['REPLACE("2009",3,2,"10")', '2010'],
    ['REPLACE("123456",1,3,"@")', '@456'],
    ['REPLACE("abc",0,1,"x")', '#VALUE!'],
    ['REPT("*-",3)', '*-*-*-'],
    ['REPT("-",0)', ''],
    ['REPT("a",-1)', '#VALUE!'],
    ['REPT("ab",20000)', '#VALUE!'],
  ]);
});

test('case, cleaning and comparison', () => {
  check([
    ['UPPER("total")', 'TOTAL'],
    ['UPPER("straße")', 'STRAßE'],
    ['LOWER("E. E. Cummings")', 'e. e. cummings'],
    ['PROPER("this is a TITLE")', 'This Is A Title'],
    ['PROPER("2-way street")', '2-Way Street'],
    ['PROPER("76BudGet")', '76Budget'],
    [`PROPER("it's")`, "It'S"],
    ['PROPER("élan vital")', 'Élan Vital'],
    ['TRIM(" First Quarter Earnings ")', 'First Quarter Earnings'],
    ['TRIM("  a   b  ")', 'a b'],
    ['TRIM("a"&CHAR(160))', 'a '],
    ['CLEAN(CHAR(9)&"Monthly report"&CHAR(10))', 'Monthly report'],
    ['EXACT("word","word")', true],
    ['EXACT("Word","word")', false],
    ['EXACT("w ord","word")', false],
    ['T("Rainfall")', 'Rainfall'],
    ['T(19)', ''],
    ['T(TRUE)', ''],
    ['T(E1)', '#N/A'],
  ], DATA);
});

test('characters: CHAR and CODE (Windows-1252), UNICHAR and UNICODE', () => {
  check([
    ['CHAR(65)', 'A'],
    ['CHAR(33)', '!'],
    ['CHAR(128)', '€'],
    ['CHAR(228)', 'ä'],
    ['CHAR(0)', '#VALUE!'],
    ['CHAR(256)', '#VALUE!'],
    ['CODE("A")', 65],
    ['CODE("Alphabet")', 65],
    ['CODE("€")', 128],
    ['CODE("ä")', 228],
    ['CODE("")', '#VALUE!'],
    ['UNICHAR(66)', 'B'],
    ['UNICHAR(32)', ' '],
    ['UNICHAR(0)', '#VALUE!'],
    ['UNICHAR(128512)', '😀'],
    ['UNICODE("B")', 66],
    ['UNICODE("😀")', 128512],
    ['UNICODE("")', '#VALUE!'],
  ]);
});

test('VALUE, NUMBERVALUE and FIXED follow the locale', () => {
  check([
    ['VALUE("$1,000")', 1000],
    ['VALUE("16:48:00")-VALUE("12:00:00")', approx(0.2)],
    ['VALUE("12%")', 0.12],
    ['VALUE("2026-01-31")', 46053],
    ['VALUE("abc")', '#VALUE!'],
    ['VALUE(TRUE)', '#VALUE!'],
    ['VALUE(A6)', 0],
    ['NUMBERVALUE("2.500,27",",",".")', 2500.27],
    ['NUMBERVALUE("3.5%")', 0.035],
    ['NUMBERVALUE(" 1 234.5 ")', 1234.5],
    ['NUMBERVALUE("12%%")', approx(0.0012)],
    ['NUMBERVALUE("")', 0],
    ['NUMBERVALUE("1.2.3")', '#VALUE!'],
    ['NUMBERVALUE("1,5",",",",")', '#VALUE!'],
    ['FIXED(1234.567,1)', '1,234.6'],
    ['FIXED(1234.567,-1)', '1,230'],
    ['FIXED(-1234.567,-1,TRUE)', '-1230'],
    ['FIXED(44.332)', '44.33'],
    ['FIXED(0.5,0)', '1'],
  ], DATA);
  check([
    ['VALUE("1.234,5")', 1234.5],
    ['NUMBERVALUE("1.234,5")', 1234.5],
    ['FIXED(1234.567,1)', '1.234,6'],
  ], {}, { locale: 'de' });
});

test('TEXTBEFORE, TEXTAFTER and TEXTSPLIT', () => {
  check([
    ['TEXTBEFORE("Red riding hood"," ")', 'Red'],
    ['TEXTBEFORE("Red riding hood"," ",-1)', 'Red riding'],
    ['TEXTBEFORE("Red riding hood","RIDING",1,1)', 'Red '],
    ['TEXTBEFORE("Red riding hood","RIDING")', '#N/A'],
    ['TEXTBEFORE("abc","x",1,0,1)', 'abc'],
    ['TEXTBEFORE("abc","x",,,,"none")', 'none'],
    ['TEXTBEFORE("a-b+c",{"+","-"})', 'a'],
    ['TEXTBEFORE("abc","b",0)', '#VALUE!'],
    ['TEXTBEFORE("abc","")', ''],
    ['TEXTAFTER("Red riding hood"," ",2)', 'hood'],
    ['TEXTAFTER("Red riding hood"," ",-1)', 'hood'],
    ['TEXTAFTER("ABC","b",1,1)', 'C'],
    ['TEXTAFTER("abc","x")', '#N/A'],
    ['TEXTAFTER("abc","x",-1,0,1)', 'abc'],
    ['TEXTAFTER({"a-b";"c-d"},"-")', [['b'], ['d']]],
    ['TEXTSPLIT("a,b;c",",",";")', [['a', 'b'], ['c', '#N/A']]],
    ['TEXTSPLIT("a,,b",",",,TRUE)', [['a', 'b']]],
    ['TEXTSPLIT("a,,b",",")', [['a', '', 'b']]],
    ['TEXTSPLIT("a b",," ")', [['a'], ['b']]],
    ['TEXTSPLIT("1-2_3",{"-","_"})', [['1', '2', '3']]],
    ['TEXTSPLIT("aXbxc","x",,,1)', [['a', 'b', 'c']]],
    ['TEXTSPLIT("a;b,c",",",";",,,"-")', [['a', '-'], ['b', 'c']]],
  ]);
});

test('TEXT and DOLLAR (number formats) work in formulas', () => {
  check([
    ['TEXT(46053,"dd.mm.yyyy")', '31.01.2026'],
    ['TEXT(0.125,"0.0%")', '12.5%'],
    ['DOLLAR(1234.567)', '$1,234.57'],
  ]);
});
