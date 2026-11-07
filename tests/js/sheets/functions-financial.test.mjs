// MatOffice Sheets – financial functions against Excel's results (Excel's own help examples), at least 9 significant digits.
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

const check = (cases, cells = {}) => { for (const [f, want] of cases) expect(evaluate(f, cells), want, f); };

test('PMT, IPMT, PPMT, FV, PV and NPER', () => {
  check([
    ['PMT(0.05/12,360,200000)', approx(-1073.6432460242797, 1e-14)],
    ['PMT(0.08/12,10,10000)', approx(-1037.032089)],
    ['PMT(0.08/12,10,10000,0,1)', approx(-1030.164327)],
    ['PMT(0.06/12,18*12,0,50000)', approx(-129.0812608)],
    ['PMT(0,10,1000)', -100],
    ['PMT(0.1,0,100)', '#NUM!'],
    ['IPMT(0.1/12,1,3*12,8000)', approx(-66.66666667)],
    ['IPMT(0.1,3,3,8000)', approx(-292.4471299)],
    ['IPMT(0.1/12,1,36,8000,0,1)', 0],
    ['IPMT(0.1,4,3,8000)', '#NUM!'],
    ['PPMT(0.1/12,1,2*12,2000)', approx(-75.62318601)],
    ['PPMT(0.08,10,10,200000)', approx(-27598.05346)],
    ['PPMT(0.1,0,3,8000)', '#NUM!'],
    ['PPMT(0.1,2,3,8000)+IPMT(0.1,2,3,8000)', approx(-3216.918429)],
    ['FV(0.06/12,10,-200,-500,1)', approx(2581.403374)],
    ['FV(0.12/12,12,-1000)', approx(12682.50301)],
    ['FV(0.11/12,35,-2000,,1)', approx(82846.24637)],
    ['FV(0,10,-100)', 1000],
    ['PV(0.08/12,12*20,500,,0)', approx(-59777.14585)],
    ['PV(0,10,-100)', 1000],
    ['NPER(0.12/12,-100,-1000,10000,1)', approx(59.6738657, 1e-8)],
    ['NPER(0.12/12,-100,-1000,10000)', approx(60.0821229, 1e-8)],
    ['NPER(0.12/12,-100,-1000)', approx(-9.57859404, 1e-8)],
    ['NPER(0,-100,1000)', 10],
    ['NPER(0.1,0,1000)', '#NUM!'],
    ['ISPMT(0.1/12,1,3*12,8000000)', approx(-64814.81481)],
    ['ISPMT(0.1,1,3,8000000)', approx(-533333.3333)],
  ]);
});

test('RATE, IRR, XIRR and MIRR find the rate', () => {
  check([
    ['RATE(4*12,-200,8000)', approx(0.00770147248820165)],
    ['RATE(4*12,-200,8000)*12', approx(0.0924176698584198)],
    ['RATE(10,-100,1000)', approx(0, 1e-9)],
    ['RATE(12,100,1000)', '#NUM!'],
    ['IRR({-70000,12000,15000,18000,21000})', approx(-0.021244848)],
    ['IRR({-70000,12000,15000,18000,21000,26000})', approx(0.086630948)],
    ['IRR({-70000,12000,15000},-0.1)', approx(-0.443506941)],
    ['IRR({1,2,3})', '#NUM!'],
    ['XIRR(A1:A5,B1:B5,0.1)', approx(0.373362535)],
    ['XIRR(A1:A5,B1:B4)', '#NUM!'],
    ['MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.12)', approx(0.126094133)],
    ['MIRR({-120000,39000,30000,21000},0.1,0.12)', approx(-0.048044655)],
    ['MIRR({1,2},0.1,0.1)', '#DIV/0!'],
  ], { A1: -10000, A2: 2750, A3: 4250, A4: 3250, A5: 2750, B1: '=DATE(2008,1,1)', B2: '=DATE(2008,3,1)', B3: '=DATE(2008,10,30)', B4: '=DATE(2009,2,15)', B5: '=DATE(2009,4,1)' });
});

test('NPV and XNPV', () => {
  check([
    ['NPV(0.1,-10000,3000,4200,6800)', approx(1188.443412)],
    ['NPV(0.08,8000,9200,10000,12000,14500)-40000', approx(1922.061555)],
    ['NPV(0.1,C1:C3)', approx(1000 / 1.1 + 2000 / 1.21)],
    ['NPV(-1,1)', '#DIV/0!'],
    ['XNPV(0.09,A1:A5,B1:B5)', approx(2086.647602)],
    ['XNPV(0.09,A1:A5,B1:B4)', '#NUM!'],
  ], { A1: -10000, A2: 2750, A3: 4250, A4: 3250, A5: 2750, B1: '=DATE(2008,1,1)', B2: '=DATE(2008,3,1)', B3: '=DATE(2008,10,30)', B4: '=DATE(2009,2,15)', B5: '=DATE(2009,4,1)', C1: 1000, C2: 'x', C3: 2000 });
});

test('CUMIPMT and CUMPRINC', () => {
  check([
    ['CUMIPMT(0.09/12,30*12,125000,13,24,0)', approx(-11135.23213)],
    ['CUMIPMT(0.09/12,30*12,125000,1,1,0)', approx(-937.5)],
    ['CUMIPMT(0.09/12,30*12,125000,13,24,2)', '#NUM!'],
    ['CUMIPMT(0,30*12,125000,13,24,0)', '#NUM!'],
    ['CUMPRINC(0.09/12,30*12,125000,13,24,0)', approx(-934.1071234)],
    ['CUMPRINC(0.09/12,30*12,125000,1,1,0)', approx(-68.27827118)],
    ['CUMPRINC(0.09/12,30*12,125000,1,360,1)', approx(-125000, 1e-12)],
  ]);
});

test('depreciation: SLN, SYD, DB and DDB', () => {
  check([
    ['SLN(30000,7500,10)', 2250],
    ['SLN(1,1,0)', '#DIV/0!'],
    ['SYD(30000,7500,10,1)', approx(4090.909091)],
    ['SYD(30000,7500,10,10)', approx(409.0909091)],
    ['SYD(1,0,10,11)', '#NUM!'],
    ['DB(1000000,100000,6,1,7)', approx(186083.3333)],
    ['DB(1000000,100000,6,2,7)', approx(259639.4167)],
    ['DB(1000000,100000,6,3,7)', approx(176814.4427)],
    ['DB(1000000,100000,6,4,7)', approx(120410.6555)],
    ['DB(1000000,100000,6,5,7)', approx(81999.6464)],
    ['DB(1000000,100000,6,6,7)', approx(55841.7592)],
    ['DB(1000000,100000,6,7,7)', approx(15845.0983)],
    ['DB(1000000,100000,6,8,7)', '#NUM!'],
    ['DDB(2400,300,10*365,1)', approx(1.315068493)],
    ['DDB(2400,300,10*12,1,2)', approx(40)],
    ['DDB(2400,300,10,1,2)', approx(480)],
    ['DDB(2400,300,10,2,1.5)', approx(306)],
    ['DDB(2400,300,10,10)', approx(22.12254898)],
    ['DDB(2400,300,10,11)', '#NUM!'],
  ]);
});

test('interest rates and prices: EFFECT, NOMINAL, PDURATION, RRI, FVSCHEDULE, DOLLARDE, DOLLARFR', () => {
  check([
    ['EFFECT(0.0525,4)', approx(0.053542667)],
    ['EFFECT(0,4)', '#NUM!'],
    ['NOMINAL(0.053543,4)', approx(0.05250032, 1e-8)],
    ['PDURATION(0.025,2000,2200)', approx(3.859866163)],
    ['PDURATION(0.025/12,1000,1200)', approx(87.60547642)],
    ['RRI(96,10000,11000)', approx(0.000992837, 1e-8)],
    ['FVSCHEDULE(1,{0.09,0.11,0.1})', approx(1.33089)],
    ['DOLLARDE(1.02,16)', 1.125],
    ['DOLLARDE(1.1,32)', 1.3125],
    ['DOLLARFR(1.125,16)', 1.02],
    ['DOLLARFR(1.125,32)', 1.04],
    ['DOLLARDE(1,0)', '#DIV/0!'],
    ['DOLLARDE(1,-1)', '#NUM!'],
  ]);
});
