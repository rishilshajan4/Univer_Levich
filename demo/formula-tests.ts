/**
 * Formula verification matrix. Each case is a SELF-CONTAINED formula (inline
 * literals / array constants — no external cell refs) plus the exact expected
 * result. The demo writes the live formula next to its expected value and a
 * tolerant PASS/CHECK status, so Univer's free formula engine is verified in
 * real time across every category (SUM/MIN/MAX/VLOOKUP/HLOOKUP and ~90 more).
 *
 * A "✗ CHECK" means the live result didn't match — usually a function Univer's
 * free engine doesn't implement (shows #NAME?), a sign/format convention diff,
 * or a genuine gap. Known risky ones are labelled.
 */
export interface FormulaTest {
  category: string;
  label: string;
  formula: string; // starts with "="
  expected: string | number | boolean;
}

export const FORMULA_TESTS: FormulaTest[] = [
  // ---------- MATH ----------
  { category: "Math", label: "SUM literals", formula: "=SUM(1,2,3,4,5)", expected: 15 },
  { category: "Math", label: "PRODUCT", formula: "=PRODUCT(2,3,4)", expected: 24 },
  { category: "Math", label: "ROUND 2 decimals", formula: "=ROUND(3.14159,2)", expected: 3.14 },
  { category: "Math", label: "ROUND negative digits", formula: "=ROUND(1234.5,-2)", expected: 1200 },
  { category: "Math", label: "ROUNDUP", formula: "=ROUNDUP(3.14159,2)", expected: 3.15 },
  { category: "Math", label: "ROUNDDOWN", formula: "=ROUNDDOWN(3.789,1)", expected: 3.7 },
  { category: "Math", label: "ABS", formula: "=ABS(-42)", expected: 42 },
  { category: "Math", label: "MOD", formula: "=MOD(17,5)", expected: 2 },
  { category: "Math", label: "POWER", formula: "=POWER(2,10)", expected: 1024 },
  { category: "Math", label: "SQRT", formula: "=SQRT(144)", expected: 12 },
  { category: "Math", label: "INT", formula: "=INT(7.89)", expected: 7 },
  { category: "Math", label: "SUMPRODUCT arrays", formula: "=SUMPRODUCT({1,2,3},{4,5,6})", expected: 32 },
  { category: "Math", label: "GCD", formula: "=GCD(24,36)", expected: 12 },
  { category: "Math", label: "LCM", formula: "=LCM(4,6)", expected: 12 },

  // ---------- STATISTICAL ----------
  { category: "Statistical", label: "AVERAGE", formula: "=AVERAGE(2,4,6,8)", expected: 5 },
  { category: "Statistical", label: "MIN", formula: "=MIN(5,3,8,1)", expected: 1 },
  { category: "Statistical", label: "MAX", formula: "=MAX(5,3,8,1)", expected: 8 },
  { category: "Statistical", label: "COUNT (numbers only)", formula: '=COUNT(1,2,"a",4)', expected: 3 },
  { category: "Statistical", label: "COUNTA (non-empty)", formula: '=COUNTA(1,"a","x",4)', expected: 4 },
  { category: "Statistical", label: "MEDIAN odd count", formula: "=MEDIAN(1,2,3,4,5)", expected: 3 },
  { category: "Statistical", label: "MEDIAN even count", formula: "=MEDIAN(1,2,3,4)", expected: 2.5 },
  { category: "Statistical", label: "MODE", formula: "=MODE(1,2,2,3,4)", expected: 2 },
  { category: "Statistical", label: "STDEV (sample)", formula: "=STDEV(2,4,6)", expected: 2 },
  { category: "Statistical", label: "VAR (sample)", formula: "=VAR(2,4,6)", expected: 4 },
  { category: "Statistical", label: "LARGE 2nd", formula: "=LARGE({3,1,4,1,5,9,2,6},2)", expected: 6 },
  { category: "Statistical", label: "SMALL 2nd", formula: "=SMALL({3,1,4,1,5,9,2,6},2)", expected: 1 },
  // RANK needs a real cell range (not an {array} constant) — see Cell-range section.

  // ---------- LOOKUP ----------
  { category: "Lookup", label: "VLOOKUP exact match", formula: '=VLOOKUP("B",{"A",1;"B",2;"C",3},2,FALSE)', expected: 2 },
  { category: "Lookup", label: "VLOOKUP approximate", formula: "=VLOOKUP(2.5,{1,10;2,20;3,30},2,TRUE)", expected: 20 },
  { category: "Lookup", label: "HLOOKUP exact match", formula: '=HLOOKUP("B",{"A","B","C";1,2,3},2,FALSE)', expected: 2 },
  { category: "Lookup", label: "INDEX row/col", formula: "=INDEX({10,20,30;40,50,60},2,3)", expected: 60 },
  { category: "Lookup", label: "MATCH exact", formula: "=MATCH(30,{10,20,30,40},0)", expected: 3 },
  { category: "Lookup", label: "INDEX+MATCH combo", formula: '=INDEX({100,200,300},1,MATCH("c",{"a","b","c"},0))', expected: 300 },
  { category: "Lookup", label: "CHOOSE", formula: '=CHOOSE(3,"a","b","c","d")', expected: "c" },
  { category: "Lookup", label: "XLOOKUP (may be unsupported)", formula: '=XLOOKUP("B",{"A";"B";"C"},{1;2;3})', expected: 2 },

  // ---------- LOGICAL ----------
  { category: "Logical", label: "IF", formula: '=IF(5>3,"yes","no")', expected: "yes" },
  { category: "Logical", label: "Nested IF", formula: '=IF(5>10,"a",IF(5>3,"b","c"))', expected: "b" },
  { category: "Logical", label: "AND", formula: "=AND(TRUE,TRUE,FALSE)", expected: false },
  { category: "Logical", label: "OR", formula: "=OR(FALSE,FALSE,TRUE)", expected: true },
  { category: "Logical", label: "NOT", formula: "=NOT(FALSE)", expected: true },
  { category: "Logical", label: "IFERROR catches div/0", formula: '=IFERROR(1/0,"err")', expected: "err" },
  { category: "Logical", label: "IFS", formula: "=IFS(FALSE,1,TRUE,2,TRUE,3)", expected: 2 },
  { category: "Logical", label: "SWITCH", formula: '=SWITCH(3,1,"a",2,"b",3,"c","d")', expected: "c" },
  { category: "Logical", label: "XOR (even TRUEs=FALSE)", formula: "=XOR(TRUE,FALSE,TRUE)", expected: false },

  // ---------- TEXT ----------
  { category: "Text", label: "CONCATENATE", formula: '=CONCATENATE("Hello"," ","World")', expected: "Hello World" },
  { category: "Text", label: "CONCAT", formula: '=CONCAT("a","b","c")', expected: "abc" },
  { category: "Text", label: "LEFT", formula: '=LEFT("Spreadsheet",6)', expected: "Spread" },
  { category: "Text", label: "RIGHT", formula: '=RIGHT("Spreadsheet",5)', expected: "sheet" },
  { category: "Text", label: "MID", formula: '=MID("Spreadsheet",7,5)', expected: "sheet" },
  { category: "Text", label: "LEN", formula: '=LEN("Univer")', expected: 6 },
  { category: "Text", label: "UPPER", formula: '=UPPER("univer")', expected: "UNIVER" },
  { category: "Text", label: "LOWER", formula: '=LOWER("UNIVER")', expected: "univer" },
  { category: "Text", label: "PROPER", formula: '=PROPER("hello world")', expected: "Hello World" },
  { category: "Text", label: "TRIM (collapse spaces)", formula: '=TRIM("  extra  spaces  ")', expected: "extra spaces" },
  { category: "Text", label: "SUBSTITUTE all", formula: '=SUBSTITUTE("a-b-c-d","-","+")', expected: "a+b+c+d" },
  { category: "Text", label: "REPLACE", formula: '=REPLACE("abcdef",2,3,"XYZ")', expected: "aXYZef" },
  { category: "Text", label: "FIND (case-sensitive)", formula: '=FIND("c","abcdef")', expected: 3 },
  { category: "Text", label: "SEARCH (case-insensitive)", formula: '=SEARCH("C","abcdef")', expected: 3 },
  { category: "Text", label: "TEXT currency format", formula: '=TEXT(1234.567,"$#,##0.00")', expected: "$1,234.57" },
  { category: "Text", label: "VALUE", formula: '=VALUE("123.45")', expected: 123.45 },
  { category: "Text", label: "REPT", formula: '=REPT("ab",3)', expected: "ababab" },
  { category: "Text", label: "TEXTJOIN", formula: '=TEXTJOIN("-",TRUE,"a","b","c")', expected: "a-b-c" },

  // ---------- DATE (serial = Excel 1900 system; compared numerically) ----------
  { category: "Date", label: "DATE serial (1900 system)", formula: "=DATE(2026,1,15)", expected: 46037 },
  { category: "Date", label: "YEAR", formula: "=YEAR(DATE(2026,1,15))", expected: 2026 },
  { category: "Date", label: "MONTH", formula: "=MONTH(DATE(2026,6,29))", expected: 6 },
  { category: "Date", label: "DAY", formula: "=DAY(DATE(2026,1,15))", expected: 15 },
  { category: "Date", label: "EOMONTH end-of-Feb (non-leap)", formula: "=DAY(EOMONTH(DATE(2026,2,15),0))", expected: 28 },
  { category: "Date", label: "EDATE clamps to month end", formula: "=DAY(EDATE(DATE(2026,1,31),1))", expected: 28 },
  { category: "Date", label: "DATEDIF days", formula: '=DATEDIF(DATE(2026,1,1),DATE(2026,12,31),"D")', expected: 364 },
  { category: "Date", label: "WEEKDAY (Thu, type1)", formula: "=WEEKDAY(DATE(2026,1,15),1)", expected: 5 },
  { category: "Date", label: "DATEVALUE serial", formula: '=DATEVALUE("2026-01-15")', expected: 46037 },

  // ---------- CONDITIONAL AGGREGATION ----------
  // SUMIF/COUNTIF/AVERAGEIF/SUMIFS/COUNTIFS/AVERAGEIFS require a real cell RANGE
  // for their criteria argument (Excel & Univer both reject {array} constants
  // there). Verified in the "Cell-range" section instead.

  // ---------- FINANCIAL (expected rounded; sign = Excel convention) ----------
  { category: "Financial", label: "PMT (10%,10yr,1000pv)", formula: "=ROUND(PMT(0.1,10,1000),2)", expected: -162.75 },
  { category: "Financial", label: "FV (10%,10yr,-100pmt)", formula: "=ROUND(FV(0.1,10,-100,0),2)", expected: 1593.74 },
  { category: "Financial", label: "PV (10%,10yr,-100pmt)", formula: "=ROUND(PV(0.1,10,-100,0),2)", expected: 614.46 },
  { category: "Financial", label: "NPV (10%,100,200,300)", formula: "=ROUND(NPV(0.1,100,200,300),2)", expected: 481.59 },
  { category: "Financial", label: "RATE (1 period)", formula: "=ROUND(RATE(1,-110,100),4)", expected: 0.1 },

  // ---------- INFORMATION ----------
  { category: "Information", label: "ISNUMBER", formula: "=ISNUMBER(42)", expected: true },
  { category: "Information", label: 'ISBLANK("") is not blank', formula: '=ISBLANK("")', expected: false },
  { category: "Information", label: "ISTEXT", formula: '=ISTEXT("hello")', expected: true },
  { category: "Information", label: "ISERROR", formula: "=ISERROR(1/0)", expected: true },
  { category: "Information", label: "ISEVEN", formula: "=ISEVEN(4)", expected: true },
  { category: "Information", label: "ISODD", formula: "=ISODD(7)", expected: true },
  { category: "Information", label: "N of TRUE", formula: "=N(TRUE)", expected: 1 },
  { category: "Information", label: "TYPE of number", formula: "=TYPE(42)", expected: 1 },

  // ---------- ARRAY / MISC (may be unsupported in older builds) ----------
  { category: "Array/Misc", label: "TRANSPOSE first value", formula: "=INDEX(TRANSPOSE({1,2,3}),1,1)", expected: 1 },
  { category: "Array/Misc", label: "SEQUENCE nth value", formula: "=INDEX(SEQUENCE(5),3)", expected: 3 },
  { category: "Array/Misc", label: "UNIQUE sum", formula: "=SUM(UNIQUE({1;2;2;3}))", expected: 6 },
];
