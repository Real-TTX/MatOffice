// MatOffice Sheets – statistical functions against Excel's results (Excel's own help examples where there are some).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluate } from './harness.mjs';

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

const DATA = {
  A1: 1, A2: 2, A3: 3, A4: 'x', A5: true, B1: 10, B2: 20, B3: 30, B4: 40, C1: 'apple', C2: 'Banana', C3: 'cherry', D1: '=""', E1: '#N/A',
};
const SAMPLE = '{1345,1301,1368,1322,1310,1370,1318,1350,1303,1299}';

test('COUNT, COUNTA and COUNTBLANK', () => {
  check([
    ['COUNT(A1:A5)', 3],
    ['COUNT(A1:A5,"3",TRUE,"x")', 5],
    ['COUNT(E1)', 0],
    ['COUNT(A:A)', 3],
    ['COUNTA(A1:A5)', 5],
    ['COUNTA(A1:A5,D1,F1)', 6],
    ['COUNTA(E1)', 1],
    ['COUNTBLANK(A1:A6)', 1],
    ['COUNTBLANK(D1:D3)', 3],
    ['COUNTBLANK(A:A)', 1048571],
  ], DATA);
});

test('COUNTIF and COUNTIFS: criteria, wildcards, blanks and whole columns', () => {
  check([
    ['COUNTIF(A1:A3,">1")', 2],
    ['COUNTIF(A1:A3,"2")', 1],
    ['COUNTIF(A1:A3,2)', 1],
    ['COUNTIF(C1:C3,"b*")', 1],
    ['COUNTIF(C1:C3,"?pple")', 1],
    ['COUNTIF(C1:C3,"<>apple")', 2],
    ['COUNTIF(C1:C3,"APPLE")', 1],
    ['COUNTIF(A1:A5,"x")', 1],
    ['COUNTIF(A1:A5,TRUE)', 1],
    ['COUNTIF(A1:A5,F1)', 0],
    ['COUNTIF(E1:E2,NA())', 1],
    ['COUNTIF(A:A,"")', 1048571],
    ['COUNTIF(A:A,"<>")', 5],
    ['COUNTIF(A1:A3,{1,3})', [[1, 1]]],
    ['COUNTIF(B1:B4,">"&A1*15)', 3],
    ['COUNTIFS(A1:A3,">1",B1:B3,"<30")', 1],
    ['COUNTIFS(A1:A3,">1",B1:B4,">1")', '#VALUE!'],
    ['COUNTIFS(A:A,"",B:B,"")', 1048571],
  ], DATA);
});

test('AVERAGEIF(S), MAXIFS and MINIFS', () => {
  check([
    ['AVERAGEIF(A1:A3,">1")', 2.5],
    ['AVERAGEIF(A1:A3,">1",B1:B3)', 25],
    ['AVERAGEIF(A1:A3,">1",B1)', 25],
    ['AVERAGEIF(A1:A3,">5")', '#DIV/0!'],
    ['AVERAGEIF(C1:C3,"*e*",B1:B3)', 20],
    ['AVERAGEIFS(B1:B3,A1:A3,">=2",C1:C3,"<>cherry")', 20],
    ['AVERAGEIFS(B1:B3,A1:A3,">9")', '#DIV/0!'],
    ['MAXIFS(B1:B3,A1:A3,">1")', 30],
    ['MINIFS(B1:B3,A1:A3,">1")', 20],
    ['MAXIFS(B1:B3,A1:A3,">9")', 0],
    ['MINIFS(B1:B4,A1:A3,">1")', '#VALUE!'],
  ], DATA);
});

test('averages, extremes, median and mode', () => {
  check([
    ['AVERAGE(A1:A5)', 2],
    ['AVERAGE(1,"2",TRUE)', approx(4 / 3)],
    ['AVERAGE(A4)', '#DIV/0!'],
    ['AVERAGE(A1:A3,E1)', '#N/A'],
    ['AVERAGEA(A1:A5)', 1.4],
    ['MAX(A1:A5)', 3],
    ['MAX(A4)', 0],
    ['MAX(-1,"-5")', -1],
    ['MAXA(A1:A5)', 3],
    ['MIN(B1:B4)', 10],
    ['MIN(A1:A5)', 1],
    ['MINA(A4:A5)', 0],
    ['MEDIAN(1,2,3,4,5)', 3],
    ['MEDIAN(1,2,3,4,5,6)', 3.5],
    ['MEDIAN(A4)', '#NUM!'],
    ['MODE({5.6,4,4,3,2,4})', 4],
    ['MODE.SNGL(1,2,2,3,3)', 2],
    ['MODE({1,2,3})', '#N/A'],
    ['MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1})', [[1], [2], [3]]],
    ['GEOMEAN(4,5,8,7,11,4,3)', approx(5.476986969)],
    ['GEOMEAN(4,9)', approx(6)],
    ['GEOMEAN(1,-1)', '#NUM!'],
    ['HARMEAN(4,5,8,7,11,4,3)', approx(5.028375962)],
    ['TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)', approx(3.777777778)],
    ['TRIMMEAN({1,2,3},1)', '#NUM!'],
  ], DATA);
});

test('spread: standard deviation, variance, deviations, skew and kurtosis', () => {
  check([
    [`STDEV.S(${SAMPLE})`, approx(27.46391572)],
    [`STDEV(${SAMPLE})`, approx(27.46391572)],
    [`STDEV.P(${SAMPLE})`, approx(26.05455814)],
    [`STDEVP(${SAMPLE})`, approx(26.05455814)],
    [`VAR.S(${SAMPLE})`, approx(754.2666667)],
    [`VAR(${SAMPLE})`, approx(754.2666667)],
    [`VAR.P(${SAMPLE})`, approx(678.84)],
    [`VARP(${SAMPLE})`, approx(678.84)],
    ['STDEVA(A1:A5)', approx(1.140175425)],
    ['STDEVPA(A1:A5)', approx(1.019803903)],
    ['VARA(A1:A5)', approx(1.3)],
    ['VARPA(A1:A5)', approx(1.04)],
    ['STDEV(1)', '#DIV/0!'],
    ['VAR.P(1)', 0],
    ['AVEDEV(4,5,6,7,5,4,3)', approx(1.020408163)],
    ['DEVSQ(4,5,8,7,11,4,3)', approx(48)],
    ['SKEW({3,4,5,2,3,4,5,6,4,7})', approx(0.359543071)],
    ['SKEW.P({3,4,5,2,3,4,5,6,4,7})', approx(0.303193339)],
    ['SKEW(1,2)', '#DIV/0!'],
    ['KURT({3,4,5,2,3,4,5,6,4,7})', approx(-0.151799637)],
    ['KURT(1,2,3)', '#DIV/0!'],
    ['STANDARDIZE(42,40,1.5)', approx(1.333333333)],
    ['STANDARDIZE(42,40,0)', '#NUM!'],
  ], DATA);
});

test('LARGE, SMALL, percentiles, quartiles and ranks', () => {
  check([
    ['LARGE({3,5,3,5,4;4,2,4,6,7},3)', 5],
    ['LARGE({3,5,3,5,4;4,2,4,6,7},7)', 4],
    ['LARGE(B1:B4,{1,2})', [[40, 30]]],
    ['LARGE(B1:B4,5)', '#NUM!'],
    ['LARGE(B1:B4,0)', '#NUM!'],
    ['SMALL({3,4,5,2,3,4,6,4,7},4)', 4],
    ['SMALL(A1:A5,1)', 1],
    ['PERCENTILE({1,3,2,4},0.3)', approx(1.9)],
    ['PERCENTILE.INC({1,3,2,4},0.3)', approx(1.9)],
    ['PERCENTILE.INC({1,2,3,4,5},0.3)', approx(2.2)],
    ['PERCENTILE.INC({1,2},1.1)', '#NUM!'],
    ['PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.25)', approx(2.5)],
    ['PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0)', '#NUM!'],
    ['PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.01)', '#NUM!'],
    ['QUARTILE({1,2,4,7,8,9,10,12},1)', 3.5],
    ['QUARTILE.INC({1,2,4,7,8,9,10,12},3)', 9.25],
    ['QUARTILE.INC({1,2,4,7,8,9,10,12},5)', '#NUM!'],
    ['QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},1)', 15],
    ['QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},3)', 43],
    ['QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},0)', '#NUM!'],
    ['PERCENTRANK({13,12,11,8,4,3,2,1,1,1},2)', 0.333],
    ['PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},4)', 0.555],
    ['PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},8)', 0.666],
    ['PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},5)', 0.583],
    ['PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},30)', '#N/A'],
    ['PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},7)', 0.7],
    ['PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43)', 0.381],
    ['PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43,1)', 0.3],
    ['RANK(3.5,{7,3.5,3.5,1,2},1)', 3],
    ['RANK(7,{7,3.5,3.5,1,2})', 1],
    ['RANK.EQ(3.5,{7,3.5,3.5,1,2})', 2],
    ['RANK.AVG(3.5,{7,3.5,3.5,1,2})', 2.5],
    ['RANK.AVG(94,{89,88,92,101,94,97,95})', 4],
    ['RANK(5,{1,2})', '#N/A'],
    ['RANK(A1:A3,A1:A3)', [[3], [2], [1]]],
  ], DATA);
});

test('two variables: correlation, covariance and the regression line', () => {
  check([
    ['CORREL({3,2,4,5,6},{9,7,12,15,17})', approx(0.997054486)],
    ['PEARSON({9,7,5,3,1},{10,6,1,5,3})', approx(0.699379, 1e-6)],
    ['RSQ({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', approx(0.05795019, 1e-8)],
    ['COVAR({3,2,4,5,6},{9,7,12,15,17})', approx(5.2)],
    ['COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})', approx(5.2)],
    ['COVARIANCE.S({2,4,8},{5,11,12})', approx(9.666666667)],
    ['SLOPE({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', approx(0.305555556)],
    ['INTERCEPT({2,3,9,1,8},{6,5,11,7,5})', approx(0.048387097)],
    ['STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})', approx(3.305719, 1e-6)],
    ['FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})', approx(10.60725308)],
    ['FORECAST.LINEAR(30,{6,7,9,15,21},{20,28,31,38,40})', approx(10.60725308)],
    ['CORREL({1,2},{1,2,3})', '#N/A'],
    ['CORREL({1,1},{1,2})', '#DIV/0!'],
    ['CORREL(A1:A4,B1:B4)', approx(1)],
  ], DATA);
});

test('distributions', () => {
  check([
    ['NORM.DIST(42,40,1.5,TRUE)', approx(0.908788780)],
    ['NORM.DIST(42,40,1.5,FALSE)', approx(0.10934005)],
    ['NORMDIST(42,40,1.5,TRUE)', approx(0.908788780)],
    ['NORM.DIST(1,0,0,TRUE)', '#NUM!'],
    ['NORM.S.DIST(1.333333,TRUE)', approx(0.908788726)],
    ['NORM.S.DIST(1.333333,FALSE)', approx(0.164010148)],
    ['NORMSDIST(1.333333)', approx(0.908788726)],
    ['NORM.S.DIST(-5,TRUE)', approx(2.866515719e-7, 1e-15)],
    ['NORM.INV(0.908789,40,1.5)', approx(42.000002, 1e-8)],
    ['NORMINV(0.908789,40,1.5)', approx(42.000002, 1e-8)],
    ['NORM.INV(0,0,1)', '#NUM!'],
    ['NORM.S.INV(0.908789)', approx(1.333334673)],
    ['NORMSINV(0.975)', approx(1.959963985)],
    ['LOGNORM.DIST(4,3.5,1.2,TRUE)', approx(0.039083556)],
    ['LOGNORM.DIST(4,3.5,1.2,FALSE)', approx(0.017617597)],
    ['LOGNORMDIST(4,3.5,1.2)', approx(0.039083556)],
    ['LOGNORM.INV(0.039084,3.5,1.2)', approx(4.0000252, 1e-7)],
    ['LOGINV(0.039084,3.5,1.2)', approx(4.0000252, 1e-7)],
    ['BINOM.DIST(6,10,0.5,FALSE)', 0.205078125],
    ['BINOM.DIST(6,10,0.5,TRUE)', approx(0.828125)],
    ['BINOMDIST(6,10,0.5,FALSE)', 0.205078125],
    ['BINOM.DIST(11,10,0.5,FALSE)', '#NUM!'],
    ['POISSON.DIST(2,5,TRUE)', approx(0.124652019)],
    ['POISSON.DIST(2,5,FALSE)', approx(0.084224337)],
    ['POISSON(2,5,FALSE)', approx(0.084224337)],
    ['EXPON.DIST(0.2,10,TRUE)', approx(0.864664717)],
    ['EXPON.DIST(0.2,10,FALSE)', approx(1.353352832)],
    ['EXPONDIST(0.2,10,TRUE)', approx(0.864664717)],
    ['EXPON.DIST(-1,1,TRUE)', '#NUM!'],
    ['WEIBULL.DIST(105,20,100,TRUE)', approx(0.929581, 1e-6)],
    ['WEIBULL.DIST(105,20,100,FALSE)', approx(0.035589, 1e-5)],
    ['WEIBULL(105,20,100,TRUE)', approx(0.929581, 1e-6)],
    ['GAMMA(2.5)', approx(1.329340388)],
    ['GAMMA(-3.75)', approx(0.267866129)],
    ['GAMMA(5)', 24],
    ['GAMMA(0)', '#NUM!'],
    ['GAMMA(-2)', '#NUM!'],
    ['GAMMALN(4)', approx(1.791759469)],
    ['GAMMALN.PRECISE(4.5)', approx(2.453736571)],
    ['GAMMALN(0)', '#NUM!'],
    ['FISHER(0.75)', approx(0.972955075)],
    ['FISHER(1)', '#NUM!'],
    ['FISHERINV(0.972955)', approx(0.75, 1e-7)],
    ['CONFIDENCE.NORM(0.05,2.5,50)', approx(0.692951912)],
    ['CONFIDENCE(0.05,2.5,50)', approx(0.692951912)],
    ['CONFIDENCE.NORM(0,2.5,50)', '#NUM!'],
    ['PHI(0.75)', approx(0.301137432)],
    ['GAUSS(2)', approx(0.477249868)],
  ]);
});

test('FREQUENCY counts into the bins and returns one more value than bins', () => {
  check([
    ['FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89})', [[1], [2], [4], [2]]],
    ['FREQUENCY({1,2,3},{3,1})', [[2], [1], [0]]],
    ['FREQUENCY(A1:A5,2)', [[2], [1]]],
    ['FREQUENCY(A1:A3,E1)', '#N/A'],
  ], DATA);
});
