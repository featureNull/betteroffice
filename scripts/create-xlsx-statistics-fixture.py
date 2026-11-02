"""Builds crates/betteroffice-xlsx/tests/fixtures/msa-statistics.xlsx.

Copies the Bosch Heft 10 (Verfahren 1) measurement system analysis template,
adds sheets of statistics formulas over its measurements and over fixed
arguments, and lets Excel compute and cache every result. The cached values
are the reference the `msa_statistics` test compares the engine against.

Needs Windows, Excel and pywin32:

    python scripts/create-xlsx-statistics-fixture.py <template.xlsx>
"""

import pathlib
import shutil
import sys

import win32com.client

ROOT = pathlib.Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "crates/betteroffice-xlsx/tests/fixtures/msa-statistics.xlsx"

Y = "V1_Messwerte"
X = "Import!$B$45:$B$94"

MEASUREMENTS = [
    f"=STEYX({Y},{X})",
    f"=PEARSON({Y},{X})",
    f"=FISHER(PEARSON({Y},{X}))",
    f"=SLOPE({Y},{X})",
    f"=INTERCEPT({Y},{X})",
    f"=RSQ({Y},{X})",
    f"=FORECAST.LINEAR(51,{Y},{X})",
    f"=STANDARDIZE(MAX({Y}),AVERAGE({Y}),STDEV.S({Y}))",
    f"=PERCENTILE.EXC({Y},0.05)",
    f"=PERCENTILE.EXC({Y},0.95)",
    f"=PERCENTILE.EXC({Y},0.01)",
    f"=QUARTILE.EXC({Y},1)",
    f"=QUARTILE.EXC({Y},3)",
    f"=PERCENTRANK.INC({Y},6.001)",
    f"=PERCENTRANK.INC({Y},6.0005,6)",
    f"=PERCENTRANK.EXC({Y},6.001,5)",
    f"=PERCENTRANK({Y},6.0015)",
    f"=PERCENTRANK.INC({Y},7)",
    f"=VARA({Y})",
    f"=VARPA({Y})",
    f"=STDEVPA({Y})",
    f"=SKEW({Y})",
    f"=SKEW.P({Y})",
    f"=KURT({Y})",
    f"=Z.TEST({Y},V1_Referenzwert)",
    f"=Z.TEST({Y},V1_Referenzwert,0.001)",
    f"=AGGREGATE(16,6,{Y},0.1)",
    f"=AGGREGATE(18,6,{Y},0.1)",
    f"=AGGREGATE(19,6,{Y},3)",
    f"=AGGREGATE(19,6,{Y},4)",
    f"=T.DIST.2T(ABS(AVERAGE({Y})-V1_Referenzwert)/(STDEV.S({Y})/SQRT(COUNT({Y}))),COUNT({Y})-1)",
    f"=CONFIDENCE.T(0.05,STDEV.S({Y}),COUNT({Y}))",
    f"=SQRT((COUNT({Y})-1)*VAR.S({Y})/CHISQ.INV.RT(0.025,COUNT({Y})-1))",
    f"=SQRT((COUNT({Y})-1)*VAR.S({Y})/CHISQ.INV(0.025,COUNT({Y})-1))",
    f"=T.INV.2T(0.05,COUNT({Y})-1)",
    f"=NORM.S.INV(0.99865)",
]

# array formulas: anchor, formula; Excel spills them
MEASUREMENT_ARRAYS = [
    ("D2", f"=LINEST({Y},{X},TRUE,TRUE)"),
    ("D8", f"=LOGEST({Y},{X},TRUE,TRUE)"),
    ("D14", f"=LINEST({Y},{X},FALSE,TRUE)"),
    ("D20", f"=LOGEST({Y},{X},FALSE,TRUE)"),
    ("D26", f"=TREND({Y},{X},{{51;52;53}})"),
    ("D30", f"=GROWTH({Y},{X},{{51;52;53}})"),
    ("D34", f"=MODE.MULT({Y})"),
    ("D40", f"=FREQUENCY({Y},{{5.999;6;6.001;6.002}})"),
]

# x1, x2 = 2 * x1, x3 = x1 squared, and y = 3 + 2 x1 plus noise
REGRESSION_DATA = [
    (1, 2, 1, 5.1),
    (2, 4, 4, 6.9),
    (3, 6, 9, 9.2),
    (4, 8, 16, 10.8),
    (5, 10, 25, 13.1),
    (6, 12, 36, 15.0),
    (7, 14, 49, 16.8),
    (8, 16, 64, 19.3),
]

# a constant predictor beside the intercept is left out: Excel then reports
# statistics that disagree with its own coefficients. an exact fit is left out
# too: its statistics are rounding noise.
REGRESSION_ARRAYS = [
    ("F2", "=LINEST(D2:D9,A2:A9,TRUE,TRUE)"),
    ("F8", "=LINEST(D2:D9,A2:B9,TRUE,TRUE)"),
    ("F14", "=LINEST(D2:D9,A2:C9,TRUE,TRUE)"),
    ("F20", "=LINEST(D2:D9,A2:C9,FALSE,TRUE)"),
    ("F26", "=LINEST(D2:D9,A2:A9,FALSE,TRUE)"),
    ("F32", "=LOGEST(D2:D9,A2:A9,TRUE,TRUE)"),
    ("F38", "=TREND(D2:D9,A2:B9,A10:B11)"),
    ("F42", "=GROWTH(D2:D9,A2:A9,{9;10})"),
    ("F46", "=MODE.MULT({1,2,2,3,3,1,4})"),
    ("F54", "=MODE.MULT({1,2,3})"),
    ("F56", "=LINEST(D2:D9,C2:C9,TRUE,TRUE)"),
    ("F62", "=TREND(D2:D9,A2:C9,A10:C11)"),
]

DISTRIBUTIONS = [
    "=TINV(0.05,49)",
    "=TINV(0.05,24)",
    "=TINV(0.05,1)",
    "=TINV(0.05,2)",
    "=TINV(0.05,1000000)",
    "=TINV(1E-10,5)",
    "=TINV(0.999999,3)",
    "=T.INV.2T(0.5,2)",
    "=T.INV(0.975,10)",
    "=T.INV(1E-12,1)",
    "=TINV(0,10)",
    "=TINV(0.05,0)",
    "=TINV(1.5,10)",
    "=TINV(1.9999,10)",
    "=TINV(2,10)",
    "=T.INV.2T(1.5,10)",
    "=T.INV.2T(1,10)",
    "=NORM.S.INV(0.5)",
    "=NORM.S.INV(0.975)",
    "=NORM.S.INV(0.99865)",
    "=NORM.S.INV(1E-10)",
    "=NORM.S.INV(1E-300)",
    "=NORM.S.INV(0.999999999999)",
    "=NORM.S.INV(0)",
    "=NORM.S.INV(1)",
    "=NORM.S.DIST(1.96,TRUE)",
    "=NORM.S.DIST(-8,TRUE)",
    "=NORM.S.DIST(-37,TRUE)",
    "=NORM.S.DIST(8,TRUE)",
    "=NORM.S.DIST(0.5,FALSE)",
    "=NORM.INV(0.99865,6,0.001)",
    "=CHISQ.INV.RT(0.05,1)",
    "=CHISQ.INV.RT(0.05,49)",
    "=CHISQ.INV.RT(0.975,49)",
    "=CHISQ.INV.RT(1E-10,10)",
    "=CHISQ.INV.RT(0.5,1000000)",
    "=CHISQ.INV.RT(0.999999,2)",
    "=CHISQ.INV(0.025,49)",
    "=CHISQ.INV.RT(0,5)",
    "=CHISQ.INV.RT(1.1,5)",
    "=CHISQ.INV.RT(0.05,0)",
    "=F.INV.RT(0.05,1,1)",
    "=F.INV.RT(0.05,5,10)",
    "=F.INV.RT(0.01,49,49)",
    "=F.INV.RT(0.5,100000,100000)",
    "=F.INV.RT(0.999,3,7)",
    "=F.INV(0.05,5,10)",
    "=F.INV.RT(0.05,0,5)",
    "=GAMMA.INV(0.5,0.1,1)",
    "=GAMMA.INV(0.99,100,2)",
    "=GAMMA.INV(1E-8,2,1)",
    "=GAMMA.INV(0.068094,9,2)",
    "=GAMMA.INV(1,2,3)",
    "=BETA.INV(0.5,0.5,0.5)",
    "=BETA.INV(0.01,2,3)",
    "=BETA.INV(0.999,10,0.5)",
    "=BETA.INV(0.685470581,8,10,1,3)",
    "=BETA.INV(0,2,3)",
    "=BETA.INV(1,2,3)",
    "=BINOM.DIST(3,10,0.3,FALSE)",
    "=BINOM.DIST(3,10,0.3,TRUE)",
    "=BINOM.DIST(50,100,0.5,FALSE)",
    "=BINOM.DIST(0,1000,0.001,TRUE)",
    "=BINOM.DIST(6,10,0.5,TRUE)",
    "=BINOM.DIST(10,10,0.5,FALSE)",
    "=BINOM.DIST(0,10,0,FALSE)",
    "=BINOM.DIST(11,10,0.3,FALSE)",
    "=BINOM.DIST(3,10,1.2,FALSE)",
    "=BINOMDIST(6,10,0.5,FALSE)",
    "=BINOM.DIST.RANGE(60,0.75,48)",
    "=BINOM.DIST.RANGE(60,0.75,45,50)",
    "=BINOM.DIST.RANGE(10,0.3,0,10)",
    "=BINOM.DIST.RANGE(10,0.3,5,4)",
    "=BINOM.INV(6,0.5,0.75)",
    "=BINOM.INV(1000,0.01,0.999)",
    "=BINOM.INV(10,0.3,0)",
    "=BINOM.INV(10,0.3,1)",
    "=BINOM.INV(10,0,0.5)",
    "=BINOM.INV(10,1,0.5)",
    "=BINOM.INV(0,0.3,0.5)",
    "=CRITBINOM(6,0.5,0.75)",
    "=POISSON.DIST(2,5,FALSE)",
    "=POISSON.DIST(2,5,TRUE)",
    "=POISSON.DIST(100,100,TRUE)",
    "=POISSON.DIST(0,0.5,FALSE)",
    "=POISSON.DIST(3,0,TRUE)",
    "=POISSON.DIST(-1,5,TRUE)",
    "=POISSON(2,5,TRUE)",
    "=HYPGEOM.DIST(1,4,8,20,TRUE)",
    "=HYPGEOM.DIST(1,4,8,20,FALSE)",
    "=HYPGEOM.DIST(4,4,8,20,TRUE)",
    "=HYPGEOM.DIST(5,4,8,20,FALSE)",
    "=HYPGEOM.DIST(5,4,8,20,TRUE)",
    "=HYPGEOM.DIST(0,15,8,20,FALSE)",
    "=HYPGEOM.DIST(0,15,8,20,TRUE)",
    "=HYPGEOM.DIST(3,15,8,20,TRUE)",
    "=HYPGEOM.DIST(-1,4,8,20,FALSE)",
    "=HYPGEOM.DIST(1,21,8,20,FALSE)",
    "=HYPGEOM.DIST(1,4,0,20,FALSE)",
    "=HYPGEOMDIST(1,4,8,20)",
    "=NEGBINOM.DIST(10,5,0.25,FALSE)",
    "=NEGBINOM.DIST(10,5,0.25,TRUE)",
    "=NEGBINOM.DIST(0,1,1,FALSE)",
    "=NEGBINOM.DIST(0,1,0,FALSE)",
    "=NEGBINOM.DIST(2,1,1,TRUE)",
    "=NEGBINOM.DIST(2,0,0.5,TRUE)",
    "=NEGBINOMDIST(10,5,0.25)",
    "=EXPON.DIST(0.2,10,TRUE)",
    "=EXPON.DIST(0.2,10,FALSE)",
    "=EXPON.DIST(1E-10,1,TRUE)",
    "=EXPONDIST(0.2,10,TRUE)",
    "=WEIBULL.DIST(105,20,100,TRUE)",
    "=WEIBULL.DIST(105,20,100,FALSE)",
    "=WEIBULL(105,20,100,TRUE)",
    "=FISHER(0.75)",
    "=FISHER(1)",
    "=FISHERINV(0.972955)",
    "=PERMUT(100,3)",
    "=PERMUT(3,2)",
    "=PERMUT(0,0)",
    "=PERMUT(-1,0)",
    "=PERMUT(5,0)",
    "=PERMUT(2,3)",
    "=PERMUTATIONA(3,2)",
    "=PERMUTATIONA(0,0)",
    "=STANDARDIZE(42,40,1.5)",
    "=STANDARDIZE(42,40,0)",
    "=Z.TEST({3,6,7,8,6,5,4,2,1,9},4)",
    "=Z.TEST({3,6,7,8,6,5,4,2,1,9},6)",
    "=ZTEST({3,6,7,8,6,5,4,2,1,9},4,2)",
    "=PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)",
    "=PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)",
    "=PROB({0,1,2,3},{0.2,0.3,0,0.5},1,3)",
    "=PROB({0,1,2,3},{0.2,0.3,0.1,0.3},1,3)",
    "=PROB({0,1,2,3},{0.2,0.3,-0.1,0.6},1,3)",
    "=PROB({0,1,2},{0.2,0.3,0.5,0.1},1,3)",
    "=PROB({0,1},{1.5,-0.5},0)",
    "=PERCENTRANK.INC({1,2,3,4,5,6,7,8,9,10},4)",
    "=PERCENTRANK.INC({1,1,3},2)",
    "=PERCENTRANK.EXC({1,1,3},2)",
    "=PERCENTRANK.INC({1,2,3,4,5,6,7,8,9,10},2/3,2)",
    "=PERCENTRANK.INC({1,2,3,4,5,6,7,8,9,10},20/3,2)",
    "=PERCENTRANK.INC({1,2,3,4,5,6,7,8,9,10},8,0)",
    "=PERCENTRANK.EXC({1,2,3,4,5,6,7,8,9,10},8)",
    "=PERCENTRANK.INC({5},5)",
    "=PERCENTILE.EXC({1,2,3,4,5,6,7,8,9},0.25)",
    "=PERCENTILE.EXC({1,2,3,4,5,6,7,8,9},0.05)",
    "=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},1)",
    "=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},0)",
    "=SKEW.P({3,4,5,2,3,4,5,6,4,7})",
    "=VARA({1,2,3})",
    "=VARPA({1,2,3})",
    "=STDEVPA({1,2,3})",
    "=STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})",
]


def write_list(sheet, formulas):
    for row, formula in enumerate(formulas, start=1):
        sheet.Cells(row, 1).Value = "'" + formula
        sheet.Cells(row, 2).Formula2 = formula


def main(template):
    FIXTURE.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template, FIXTURE)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(FIXTURE))
        last = book.Worksheets(book.Worksheets.Count)

        measurements = book.Worksheets.Add(After=last)
        measurements.Name = "Statistik"
        write_list(measurements, MEASUREMENTS)
        for anchor, formula in MEASUREMENT_ARRAYS:
            measurements.Range(anchor).Formula2 = formula

        regression = book.Worksheets.Add(After=measurements)
        regression.Name = "Regression"
        for row, values in enumerate(REGRESSION_DATA, start=2):
            for col, value in enumerate(values, start=1):
                regression.Cells(row, col).Value = value
        for row, values in enumerate([(9, 18, 81), (10, 20, 100)], start=10):
            for col, value in enumerate(values, start=1):
                regression.Cells(row, col).Value = value
        for anchor, formula in REGRESSION_ARRAYS:
            regression.Range(anchor).Formula2 = formula

        distributions = book.Worksheets.Add(After=regression)
        distributions.Name = "Verteilungen"
        write_list(distributions, DISTRIBUTIONS)

        excel.CalculateFull()
        book.Save()
        book.Close(False)
    finally:
        excel.Quit()
    print(f"wrote {FIXTURE.relative_to(ROOT)}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(pathlib.Path(sys.argv[1]).resolve())
