// MatOffice Sheets – math & trigonometry functions against Excel's results (Excel's own help examples where there are some).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluate, workbook } from './harness.mjs';

/** A number expected within a relative tolerance. */
const approx = (value, tol = 1e-9) => ({ approx: value, tol });

function expect(got, want, label) {
  if (want && typeof want === 'object' && 'approx' in want) {
    assert.equal(typeof got, 'number', `${label}: ${JSON.stringify(got)} is not a number`);
    assert.ok(Math.abs(got - want.approx) <= want.tol * Math.max(1, Math.abs(want.approx)), `${label}: ${got} is not ${want.approx}`);
  } else if (Array.isArray(want)) {
    assert.ok(Array.isArray(got) && got.length === want.length, `${label}: ${JSON.stringify(got)} is not ${JSON.stringify(want)}`);
    want.forEach((w, i) => expect(got[i], w, `${label}[${i}]`));
  } else {
    assert.deepEqual(got, want, label);
  }
}

const check = (cases, cells = {}, options = {}) => { for (const [f, want] of cases) expect(evaluate(f, cells, options), want, f); };

const DATA = { A1: 1, A2: 2, A3: 3, A4: 'x', A5: true, B1: 10, B2: 20, B3: 30, B4: 40, C1: 'apple', C2: 'Banana', C3: 'cherry', E1: '#N/A' };

test('SUM, PRODUCT and SUMSQ count the numbers of ranges and convert typed arguments', () => {
  check([
    ['SUM(1,2,3)', 6],
    ['SUM(A1:A5)', 6],
    ['SUM("3",TRUE)', 4],
    ['SUM(A1:A3,E1)', '#N/A'],
    ['SUM("x")', '#VALUE!'],
    ['SUM(A:A)', 6],
    ['SUM(A1:B4)', 106],
    ['PRODUCT(2,3,4)', 24],
    ['PRODUCT(A1:A5)', 6],
    ['PRODUCT(A4)', 0],
    ['SUMSQ(3,4)', 25],
    ['SUMSQ(A1:A3)', 14],
  ], DATA);
});

test('SUMIF and SUMIFS use criteria, wildcards and the size of the criteria range', () => {
  check([
    ['SUMIF(A1:A3,">1",B1:B3)', 50],
    ['SUMIF(A1:A3,">1")', 5],
    ['SUMIF(A1:A3,">1",B1)', 50],
    ['SUMIF(A:A,">=2",B:B)', 50],
    ['SUMIF(C1:C3,"*an*",B1:B3)', 20],
    ['SUMIF(C1:C3,"<>apple",B1:B3)', 50],
    ['SUMIF(A1:A3,2,B1:B3)', 20],
    ['SUMIF(A1:A3,F1,B1:B3)', 0],
    ['SUMIF(A1:A3,{1,3},B1:B3)', [[10, 30]]],
    ['SUMIF(A1:A3,">0",E1:E3)', '#N/A'],
    ['SUMIFS(B1:B4,A1:A4,">=2")', 50],
    ['SUMIFS(B1:B3,A1:A3,">1",C1:C3,"c*")', 30],
    ['SUMIFS(B:B,A:A,">1",C:C,"<>")', 50],
    ['SUMIFS(B1:B4,A1:A3,">1")', '#VALUE!'],
  ], DATA);
});

test('SUMPRODUCT and the sums of pairs', () => {
  check([
    ['SUMPRODUCT(A1:A3,B1:B3)', 140],
    ['SUMPRODUCT((A1:A3>1)*B1:B3)', 50],
    ['SUMPRODUCT(A1:A5,B1:B5)', 140],
    ['SUMPRODUCT(A1:A3,B1:B4)', '#VALUE!'],
    ['SUMPRODUCT({1,2;3,4})', 10],
    ['SUMPRODUCT(A1:A3,E1:E3)', '#N/A'],
    ['SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', -55],
    ['SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', 521],
    ['SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', 79],
    ['SUMXMY2({1,2},{1,2,3})', '#N/A'],
    ['SERIESSUM(PI()/4,0,2,{1,-0.5,0.041666667,-0.001388889})', approx(0.707103215, 1e-8)],
  ], DATA);
});

test('rounding like Excel: halves away from zero on the 15-digit value', () => {
  check([
    ['ROUND(2.15,1)', 2.2],
    ['ROUND(2.149,1)', 2.1],
    ['ROUND(-1.475,2)', -1.48],
    ['ROUND(21.5,-1)', 20],
    ['ROUND(626.3,-3)', 1000],
    ['ROUND(1.98,-1)', 0],
    ['ROUND(-50.55,-2)', -100],
    ['ROUND(2.675,2)', 2.68],
    ['ROUND(-2.5,0)', -3],
    ['ROUND(1234.5678,-2)', 1200],
    ['ROUND("x",1)', '#VALUE!'],
    ['ROUNDUP(3.2,0)', 4],
    ['ROUNDUP(76.9,0)', 77],
    ['ROUNDUP(3.14159,3)', 3.142],
    ['ROUNDUP(-3.14159,1)', -3.2],
    ['ROUNDUP(31415.92654,-2)', 31500],
    ['ROUNDUP(0.1+0.2,1)', 0.3],
    ['ROUNDDOWN(3.2,0)', 3],
    ['ROUNDDOWN(76.9,0)', 76],
    ['ROUNDDOWN(3.14159,3)', 3.141],
    ['ROUNDDOWN(-3.14159,1)', -3.1],
    ['ROUNDDOWN(31415.92654,-2)', 31400],
    ['TRUNC(8.9)', 8],
    ['TRUNC(-8.9)', -8],
    ['TRUNC(0.45,1)', 0.4],
    ['INT(8.9)', 8],
    ['INT(-8.9)', -9],
    ['MROUND(10,3)', 9],
    ['MROUND(-10,-3)', -9],
    ['MROUND(1.3,0.2)', 1.4],
    ['MROUND(5,-2)', '#NUM!'],
    ['MROUND(5,0)', 0],
  ]);
});

test('CEILING, FLOOR and their variants', () => {
  check([
    ['CEILING(2.5,1)', 3],
    ['CEILING(-2.5,-2)', -4],
    ['CEILING(-2.5,2)', -2],
    ['CEILING(1.5,0.1)', 1.5],
    ['CEILING(0.234,0.01)', 0.24],
    ['CEILING(1.1,0.1)', 1.1],
    ['CEILING(4.42,0.05)', 4.45],
    ['CEILING(2.5,-2)', '#NUM!'],
    ['CEILING(2.5,0)', 0],
    ['FLOOR(3.7,2)', 2],
    ['FLOOR(-2.5,-2)', -2],
    ['FLOOR(-2.5,2)', -4],
    ['FLOOR(2.5,-2)', '#NUM!'],
    ['FLOOR(1.58,0.1)', 1.5],
    ['FLOOR(0.234,0.01)', 0.23],
    ['FLOOR(5,0)', '#DIV/0!'],
    ['CEILING.MATH(24.3,5)', 25],
    ['CEILING.MATH(6.7)', 7],
    ['CEILING.MATH(-8.1,2)', -8],
    ['CEILING.MATH(-5.5,2,-1)', -6],
    ['FLOOR.MATH(24.3,5)', 20],
    ['FLOOR.MATH(6.7)', 6],
    ['FLOOR.MATH(-8.1,2)', -10],
    ['FLOOR.MATH(-5.5,2,-1)', -4],
    ['CEILING.PRECISE(4.3)', 5],
    ['CEILING.PRECISE(-4.3)', -4],
    ['CEILING.PRECISE(4.3,-2)', 6],
    ['ISO.CEILING(-4.3,2)', -4],
    ['FLOOR.PRECISE(3.2)', 3],
    ['FLOOR.PRECISE(-3.2,-1)', -4],
    ['EVEN(1.5)', 2],
    ['EVEN(3)', 4],
    ['EVEN(-1)', -2],
    ['EVEN(0)', 0],
    ['ODD(1.5)', 3],
    ['ODD(3)', 3],
    ['ODD(2)', 3],
    ['ODD(-1)', -1],
    ['ODD(-2)', -3],
    ['ODD(0)', 1],
  ]);
});

test('arithmetic functions and their errors', () => {
  check([
    ['ABS(-4)', 4],
    ['SIGN(-0.5)', -1],
    ['SIGN(0)', 0],
    ['SQRT(16)', 4],
    ['SQRT(-16)', '#NUM!'],
    ['SQRTPI(1)', approx(1.772453851)],
    ['SQRTPI(2)', approx(2.506628275)],
    ['EXP(1)', approx(2.718281828)],
    ['LN(86)', approx(4.454347296)],
    ['LN(0)', '#NUM!'],
    ['LOG10(86)', approx(1.934498451)],
    ['LOG10(1E5)', 5],
    ['LOG(10)', 1],
    ['LOG(8,2)', 3],
    ['LOG(86,2.7182818)', approx(4.4543473, 1e-7)],
    ['LOG(10,1)', '#DIV/0!'],
    ['LOG(-1)', '#NUM!'],
    ['POWER(5,2)', 25],
    ['POWER(98.6,3.2)', approx(2401077.222)],
    ['POWER(4,5/4)', approx(5.656854249)],
    ['POWER(0,0)', '#NUM!'],
    ['POWER(0,-1)', '#DIV/0!'],
    ['POWER(-8,1/3)', '#NUM!'],
    ['MOD(3,2)', 1],
    ['MOD(-3,2)', 1],
    ['MOD(3,-2)', -1],
    ['MOD(-3,-2)', -1],
    ['MOD(1,0)', '#DIV/0!'],
    ['QUOTIENT(5,2)', 2],
    ['QUOTIENT(4.5,3.1)', 1],
    ['QUOTIENT(-10,3)', -3],
    ['QUOTIENT(1,0)', '#DIV/0!'],
    ['PI()', approx(3.14159265358979)],
    ['FACT(5)', 120],
    ['FACT(1.9)', 1],
    ['FACT(0)', 1],
    ['FACT(-1)', '#NUM!'],
    ['FACT(171)', '#NUM!'],
    ['FACTDOUBLE(6)', 48],
    ['FACTDOUBLE(7)', 105],
    ['FACTDOUBLE(-1)', 1],
    ['COMBIN(8,2)', 28],
    ['COMBIN(4,5)', '#NUM!'],
    ['COMBINA(4,3)', 20],
    ['COMBINA(10,3)', 220],
    ['PERMUT(100,3)', 970200],
    ['PERMUT(3,2)', 6],
    ['PERMUTATIONA(3,2)', 9],
    ['PERMUTATIONA(2,2)', 4],
    ['MULTINOMIAL(2,3,4)', 1260],
    ['GCD(5,2)', 1],
    ['GCD(24,36)', 12],
    ['GCD(5,0)', 5],
    ['GCD(-1,2)', '#NUM!'],
    ['LCM(5,2)', 10],
    ['LCM(24,36)', 72],
    ['LCM(0,5)', 0],
  ]);
  for (let i = 0; i < 5; i++) {
    const r = evaluate('RAND()');
    assert.ok(r >= 0 && r < 1);
    const d = evaluate('RANDBETWEEN(1,6)');
    assert.ok(Number.isInteger(d) && d >= 1 && d <= 6);
  }
  assert.equal(evaluate('RANDBETWEEN(5,1)'), '#NUM!');
});

test('trigonometry in radians', () => {
  check([
    ['SIN(PI()/2)', 1],
    ['SIN(PI())', approx(1.22465e-16, 1e-20)],
    ['SIN(30*PI()/180)', approx(0.5)],
    ['COS(0)', 1],
    ['COS(60*PI()/180)', approx(0.5)],
    ['TAN(PI()/4)', approx(1)],
    ['ASIN(1)', approx(Math.PI / 2)],
    ['ASIN(2)', '#NUM!'],
    ['ACOS(-0.5)', approx(2.094395102)],
    ['ACOS(1.5)', '#NUM!'],
    ['ATAN(1)', approx(0.785398163)],
    ['ATAN2(1,1)', approx(0.785398163)],
    ['ATAN2(-1,-1)', approx(-2.35619449)],
    ['ATAN2(0,0)', '#DIV/0!'],
    ['SINH(1)', approx(1.175201194)],
    ['COSH(4)', approx(27.30823284)],
    ['TANH(0.5)', approx(0.462117157)],
    ['ASINH(-2.5)', approx(-1.647231146)],
    ['ACOSH(10)', approx(2.993222846)],
    ['ACOSH(0.5)', '#NUM!'],
    ['ATANH(0.76159416)', approx(1.00000001, 1e-8)],
    ['ATANH(1)', '#NUM!'],
    ['DEGREES(PI())', 180],
    ['RADIANS(270)', approx(4.71238898)],
    ['COT(30)', approx(-0.156119952)],
    ['COT(0)', '#DIV/0!'],
    ['CSC(15)', approx(1.537780562)],
    ['SEC(45)', approx(1.903594407)],
    ['COTH(2)', approx(1.037314721)],
    ['CSCH(1.5)', approx(0.469642441)],
    ['SECH(1)', approx(0.648054274)],
    ['ACOT(2)', approx(0.463647609)],
    ['ACOTH(6)', approx(0.168236118)],
    ['ACOTH(0.5)', '#NUM!'],
  ]);
});

test('SUBTOTAL and AGGREGATE: the functions, hidden rows and nested subtotals', () => {
  check([
    ['SUBTOTAL(9,A1:A3)', 6],
    ['SUBTOTAL(1,A1:A3)', 2],
    ['SUBTOTAL(2,A1:A5)', 3],
    ['SUBTOTAL(3,A1:A5)', 5],
    ['SUBTOTAL(4,B1:B4)', 40],
    ['SUBTOTAL(5,B1:B4)', 10],
    ['SUBTOTAL(6,A1:A3)', 6],
    ['SUBTOTAL(7,A1:A3)', 1],
    ['SUBTOTAL(8,A1:A3)', approx(0.816496581)],
    ['SUBTOTAL(10,A1:A3)', 1],
    ['SUBTOTAL(11,A1:A3)', approx(0.666666667)],
    ['SUBTOTAL(109,A1:A3,B1:B4)', 106],
    ['SUBTOTAL(12,A1:A3)', '#VALUE!'],
    ['SUBTOTAL(9,A1:A3,E1)', '#N/A'],
    ['AGGREGATE(14,6,{1,#N/A,3},1)', 3],
    ['AGGREGATE(15,6,{5,#DIV/0!,2},1)', 2],
    ['AGGREGATE(9,6,A1:A3,E1)', 6],
    ['AGGREGATE(9,4,A1:A3,E1)', '#N/A'],
    ['AGGREGATE(12,0,{1,2,3,4})', 2.5],
    ['AGGREGATE(13,0,{1,2,2,3})', 2],
    ['AGGREGATE(16,0,{1,2,3,4},0.5)', 2.5],
    ['AGGREGATE(17,0,{1,2,3,4},1)', 1.75],
    ['AGGREGATE(18,0,{1,2,3,6,6,6,7,8,9},0.25)', 2.5],
    ['AGGREGATE(19,0,{6,7,15,36,39,40,41,42,43,47,49},1)', 15],
    ['AGGREGATE(20,0,A1:A3)', '#VALUE!'],
    ['AGGREGATE(14,0,{1,2,3})', '#VALUE!'],
    ['AGGREGATE(14,6,A1:A3/(B1:B3>10),1)', 3],
  ], DATA);

  const wb = workbook({
    Sheet1: {
      A1: 1, A2: 2, A3: '=SUBTOTAL(9,A1:A2)', A4: 4,
      B1: '=SUBTOTAL(9,A1:A4)', B2: '=SUBTOTAL(109,A1:A4)', B3: '=SUM(A1:A4)', B4: '=AGGREGATE(9,0,A1:A4)', B5: '=AGGREGATE(9,4,A1:A4)',
      B6: '=AGGREGATE(9,5,A1:A4)', B7: '=SUBTOTAL(103,A1:A4)', B8: '=AGGREGATE(9,1,A1:A4)',
    },
  });
  wb.host.rowHidden = (sheet, row) => row === 1; // row 2 is hidden
  wb.calc.rebuild();
  assert.equal(wb.get('B1'), 7); // the nested SUBTOTAL in A3 is left out
  assert.equal(wb.get('B2'), 5); // and the hidden row 2 too
  assert.equal(wb.get('B3'), 10);
  assert.equal(wb.get('B4'), 7);
  assert.equal(wb.get('B5'), 10);
  assert.equal(wb.get('B6'), 8);
  assert.equal(wb.get('B7'), 2);
  assert.equal(wb.get('B8'), 5);
});

test('ROMAN, ARABIC, BASE and DECIMAL', () => {
  check([
    ['ROMAN(499)', 'CDXCIX'],
    ['ROMAN(499,1)', 'LDVLIV'],
    ['ROMAN(499,2)', 'XDIX'],
    ['ROMAN(499,3)', 'VDIV'],
    ['ROMAN(499,4)', 'ID'],
    ['ROMAN(1999,TRUE)', 'MCMXCIX'],
    ['ROMAN(1999,FALSE)', 'MIM'],
    ['ROMAN(2026)', 'MMXXVI'],
    ['ROMAN(0)', ''],
    ['ROMAN(4000)', '#VALUE!'],
    ['ROMAN(-1)', '#VALUE!'],
    ['ARABIC("LVII")', 57],
    ['ARABIC("mcmxii")', 1912],
    ['ARABIC("")', 0],
    ['ARABIC("-MMXI")', -2011],
    ['ARABIC("ABC")', '#VALUE!'],
    ['BASE(7,2)', '111'],
    ['BASE(100,16)', '64'],
    ['BASE(15,2,10)', '0000001111'],
    ['BASE(-1,2)', '#NUM!'],
    ['BASE(10,37)', '#NUM!'],
    ['DECIMAL("FF",16)', 255],
    ['DECIMAL(111,2)', 7],
    ['DECIMAL("zap",36)', 45745],
    ['DECIMAL("12",2)', '#NUM!'],
  ]);
});

test('matrices: MMULT, MDETERM, MINVERSE and MUNIT', () => {
  check([
    ['MMULT({1,3;7,2},{2,0;0,2})', [[2, 6], [14, 4]]],
    ['MMULT({1,2;3,4},{5;6})', [[17], [39]]],
    ['MMULT({1,2},{1,2})', '#VALUE!'],
    ['MMULT({1,"a"},{1;2})', '#VALUE!'],
    ['MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})', 88],
    ['MDETERM({3,6;1,1})', -3],
    ['MDETERM({1,2,3;4,5,6})', '#VALUE!'],
    ['MINVERSE({4,-1;2,0})', [[0, 0.5], [-1, 2]]],
    ['MINVERSE({-1,0;2,1})', [[-1, 0], [2, 1]]],
    ['MINVERSE({1,2;2,4})', '#NUM!'],
    ['MUNIT(3)', [[1, 0, 0], [0, 1, 0], [0, 0, 1]]],
    ['MUNIT(0)', '#VALUE!'],
  ]);
});
