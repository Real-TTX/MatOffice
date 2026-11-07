// MatOffice Sheets – date and time functions against Excel's results. "Now" is fixed by the harness: 2026-01-15 12:00 (serial 46037.5).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluate } from './harness.mjs';
import { serial } from '../../../src/MatOffice/wwwroot/js/sheets/engine/dates.js';

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
const d = (y, m, day) => serial(y, m, day);

test('DATE and TIME overflow like Excel and refuse what is out of range', () => {
  check([
    ['DATE(2026,1,31)', 46053],
    ['DATE(2026,14,1)', d(2027, 2, 1)],
    ['DATE(2026,1,0)', d(2025, 12, 31)],
    ['DATE(2026,0,1)', d(2025, 12, 1)],
    ['DATE(2008,-3,2)', d(2007, 9, 2)],
    ['DATE(2008,1,35)', d(2008, 2, 4)],
    ['DATE(108,1,2)', d(2008, 1, 2)],
    ['DATE(1900,1,1)', 1],
    ['DATE(1900,2,29)', 60],
    ['DATE(1900,3,0)', 60],
    ['DATE(1900,1,61)', 61],
    ['DATE(9999,12,31)', 2958465],
    ['DATE(9999,12,32)', '#NUM!'],
    ['DATE(-1,1,1)', '#NUM!'],
    ['DATE(10000,1,1)', '#NUM!'],
    ['DATE(2026.9,1.9,1.9)', 46023],
    ['TIME(12,0,0)', 0.5],
    ['TIME(16,48,10)', approx(0.700115741)],
    ['TIME(27,0,0)', 0.125],
    ['TIME(0,750,0)', approx(0.520833333)],
    ['TIME(1,-1,0)', approx(59 / 1440)],
    ['TIME(0,-1,0)', '#NUM!'],
  ]);
});

test('TODAY and NOW read the clock of the workbook', () => {
  check([
    ['TODAY()', 46037],
    ['NOW()', 46037.5],
    ['YEAR(TODAY())', 2026],
    ['TODAY()-WEEKDAY(TODAY(),2)+1', 46034],
    ['EOMONTH(TODAY(),0)+1', d(2026, 2, 1)],
  ]);
});

test('the parts of a date and a time', () => {
  check([
    ['YEAR(46053)', 2026],
    ['YEAR("2026-01-31")', 2026],
    ['MONTH(46053)', 1],
    ['DAY(46053)', 31],
    ['YEAR(0)', 1900],
    ['MONTH(0)', 1],
    ['DAY(0)', 0],
    ['MONTH(60)', 2],
    ['DAY(60)', 29],
    ['YEAR(-1)', '#NUM!'],
    ['YEAR("x")', '#VALUE!'],
    ['YEAR(2958466)', '#NUM!'],
    ['HOUR(0.75)', 18],
    ['HOUR(46037.5)', 12],
    ['MINUTE("12:45:00")', 45],
    ['SECOND("4:48:18 PM")', 18],
    ['SECOND(TIME(0,0,30)+0.6/86400)', 31],
    ['HOUR(-1)', '#NUM!'],
  ]);
});

test('WEEKDAY, WEEKNUM and ISOWEEKNUM with their return types', () => {
  check([
    ['WEEKDAY(DATE(2008,2,14))', 5],
    ['WEEKDAY(DATE(2008,2,14),2)', 4],
    ['WEEKDAY(DATE(2008,2,14),3)', 3],
    ['WEEKDAY(DATE(2026,1,31))', 7],
    ['WEEKDAY(DATE(2026,2,1))', 1],
    ['WEEKDAY(DATE(2026,2,1),2)', 7],
    ['WEEKDAY(DATE(2026,2,1),11)', 7],
    ['WEEKDAY(DATE(2026,2,1),12)', 6],
    ['WEEKDAY(DATE(2026,2,1),17)', 1],
    ['WEEKDAY(DATE(2026,2,1),4)', '#NUM!'],
    ['WEEKDAY(0)', 7],
    ['WEEKNUM(DATE(2012,3,9))', 10],
    ['WEEKNUM(DATE(2012,3,9),2)', 11],
    ['WEEKNUM(DATE(2026,1,4))', 2],
    ['WEEKNUM(DATE(2026,1,4),2)', 1],
    ['WEEKNUM(DATE(2026,1,4),21)', 1],
    ['WEEKNUM(DATE(2027,1,1),21)', 53],
    ['WEEKNUM(DATE(2026,1,4),3)', '#NUM!'],
    ['ISOWEEKNUM(DATE(2012,3,9))', 10],
    ['ISOWEEKNUM(DATE(2026,12,31))', 53],
    ['ISOWEEKNUM(DATE(2027,1,1))', 53],
    ['ISOWEEKNUM(DATE(2024,12,30))', 1],
  ]);
});

test('EDATE and EOMONTH keep to the end of the month', () => {
  check([
    ['EDATE(DATE(2011,1,15),1)', d(2011, 2, 15)],
    ['EDATE(DATE(2011,1,15),-1)', d(2010, 12, 15)],
    ['EDATE(DATE(2011,1,15),2)', d(2011, 3, 15)],
    ['EDATE(DATE(2026,1,31),1)', d(2026, 2, 28)],
    ['EDATE(DATE(2024,2,29),12)', d(2025, 2, 28)],
    ['EDATE(DATE(2026,1,31)+0.75,1)', d(2026, 2, 28)],
    ['EDATE(DATE(2026,1,15),1.9)', d(2026, 2, 15)],
    ['EDATE(-1,1)', '#NUM!'],
    ['EDATE(DATE(1900,1,15),-1)', '#NUM!'],
    ['EOMONTH(DATE(2011,1,1),1)', d(2011, 2, 28)],
    ['EOMONTH(DATE(2011,1,1),-3)', d(2010, 10, 31)],
    ['EOMONTH(DATE(2024,1,31),1)', d(2024, 2, 29)],
    ['EOMONTH(DATE(2026,12,15),0)', d(2026, 12, 31)],
  ]);
});

test('DATEDIF, DAYS, DAYS360 and YEARFRAC', () => {
  check([
    ['DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"Y")', 2],
    ['DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"D")', 440],
    ['DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YD")', 75],
    ['DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"M")', 14],
    ['DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YM")', 2],
    ['DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"MD")', 14],
    ['DATEDIF(DATE(2020,2,29),DATE(2021,2,28),"Y")', 0],
    ['DATEDIF(DATE(2026,1,15),DATE(2025,1,1),"D")', '#NUM!'],
    ['DATEDIF(DATE(2025,1,1),DATE(2026,1,15),"X")', '#NUM!'],
    ['DAYS(DATE(2021,3,15),DATE(2021,2,1))', 42],
    ['DAYS("2021-03-15","2021-02-01")', 42],
    ['DAYS(DATE(2021,2,1),DATE(2021,3,15))', -42],
    ['DAYS360(DATE(2011,1,1),DATE(2011,12,31))', 360],
    ['DAYS360(DATE(2011,1,30),DATE(2011,12,31))', 330],
    ['DAYS360(DATE(2011,1,31),DATE(2011,12,31),TRUE)', 330],
    ['DAYS360(DATE(2011,2,28),DATE(2011,3,31))', 30],
    ['DAYS360(DATE(2011,2,28),DATE(2011,3,31),TRUE)', 32],
    ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))', approx(0.58055556, 1e-8)],
    ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)', approx(0.57650273, 1e-8)],
    ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),2)', approx(211 / 360)],
    ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)', approx(0.57808219, 1e-8)],
    ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),4)', approx(209 / 360)],
    ['YEARFRAC(DATE(2012,7,30),DATE(2012,1,1))', approx(0.58055556, 1e-8)],
    ['YEARFRAC(DATE(2010,1,1),DATE(2013,7,1),1)', approx(1277 / 365.25)],
    ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),5)', '#NUM!'],
  ]);
});

test('working days: NETWORKDAYS(.INTL) and WORKDAY(.INTL) with weekends and holidays', () => {
  const holidays = { H1: '=DATE(2012,11,22)', H2: '=DATE(2012,12,4)', H3: '=DATE(2013,1,21)', H4: '=DATE(2006,1,2)', H5: '=DATE(2006,1,16)' };
  check([
    ['NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))', 110],
    ['NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),H1)', 109],
    ['NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),H1:H3)', 107],
    ['NETWORKDAYS(DATE(2026,1,31),DATE(2026,1,1))', -22],
    ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))', 22],
    ['NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))', -21],
    ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,H4:H5)', 22],
    ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0010001",H4:H5)', 20],
    ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"1111111")', '#VALUE!'],
    ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),8)', '#NUM!'],
    ['WORKDAY(DATE(2008,10,1),151)', d(2009, 4, 30)],
    ['WORKDAY(DATE(2008,10,1),151,H6:H8)', d(2009, 5, 5)],
    ['WORKDAY(DATE(2026,1,30),1)', d(2026, 2, 2)],
    ['WORKDAY(DATE(2026,2,2),-1)', d(2026, 1, 30)],
    ['WORKDAY(DATE(2026,1,31),0)', d(2026, 1, 31)],
    ['WORKDAY.INTL(DATE(2012,1,1),30,0)', '#NUM!'],
    ['WORKDAY.INTL(DATE(2012,1,1),90,11)', d(2012, 4, 14)],
    ['WORKDAY.INTL(DATE(2012,1,1),30,17)', 40944],
  ], { ...holidays, H6: '=DATE(2008,11,26)', H7: '=DATE(2008,12,4)', H8: '=DATE(2009,1,21)' });
});

test('DATEVALUE and TIMEVALUE read typed dates and times in the locale', () => {
  check([
    ['DATEVALUE("8/22/2008")', 39682],
    ['DATEVALUE("2008-08-22")', 39682],
    ['DATEVALUE("22-MAY-2011")', 40685],
    ['DATEVALUE("2026-01-15 12:00")', 46037],
    ['DATEVALUE("1/31")', 46053],
    ['DATEVALUE(46037)', '#VALUE!'],
    ['DATEVALUE("x")', '#VALUE!'],
    ['TIMEVALUE("2:24 AM")', approx(0.1)],
    ['TIMEVALUE("22-Aug-2011 6:35 AM")', approx(0.274305556)],
    ['TIMEVALUE("36:00")', 0.5],
    ['TIMEVALUE("2026-01-15")', 0],
    ['TIMEVALUE("x")', '#VALUE!'],
  ]);
  check([['DATEVALUE("31.01.2026")', 46053], ['DATEVALUE("31.1.")', 46053]], {}, { locale: 'de' });
});
