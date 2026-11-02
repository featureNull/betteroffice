//! the statistics functions added beside the distribution families: discrete
//! distributions, exclusive percentiles, percent ranks, the A-variance forms,
//! regression helpers and MODE.MULT/LOGEST/GROWTH. expected values are the
//! Microsoft Learn examples and edge cases as Excel 365 computes them (see
//! `scripts/create-xlsx-statistics-fixture.py`).

use xlsx_calc::{EvalContext, Value, evaluate, evaluate_array, parse_formula};
use xlsx_model::{CellValue, ErrorValue, SheetId, Workbook};

fn e(v: ErrorValue) -> CellValue {
    CellValue::Error { value: v }
}

fn n(v: f64) -> CellValue {
    CellValue::Number { value: v }
}

fn workbook() -> Workbook {
    let mut wb = Workbook::default();
    wb.sheets.push(xlsx_model::Sheet::new("Sheet1"));
    wb
}

fn eval(src: &str) -> CellValue {
    let wb = workbook();
    let expr = parse_formula(src).expect("parse");
    evaluate(&expr, &EvalContext::new(&wb, SheetId(0)))
}

fn block(src: &str) -> (usize, usize, Vec<CellValue>) {
    let wb = workbook();
    let expr = parse_formula(src).expect("parse");
    match evaluate_array(&expr, &EvalContext::new(&wb, SheetId(0))) {
        Value::Array(array) => (array.rows(), array.cols(), array.cells().to_vec()),
        other => (1, 1, vec![other.into_scalar()]),
    }
}

fn close(got: &CellValue, want: f64, src: &str) {
    match got {
        CellValue::Number { value } => assert!(
            (value - want).abs() <= 1e-12 * want.abs().max(1e-300) || value == &want,
            "{src}: {value} != {want}"
        ),
        other => panic!("{src}: expected {want}, got {other:?}"),
    }
}

fn approx(cases: &[(&str, f64)]) {
    for (src, want) in cases {
        close(&eval(src), *want, src);
    }
}

fn check(cases: &[(&str, CellValue)]) {
    for (src, want) in cases {
        assert_eq!(eval(src), *want, "{src}");
    }
}

#[test]
fn binomial_family() {
    approx(&[
        ("BINOM.DIST(6,10,0.5,FALSE)", 0.205078125),
        ("BINOMDIST(6,10,0.5,FALSE)", 0.205078125),
        ("BINOM.DIST(6,10,0.5,TRUE)", 0.828125),
        ("BINOM.DIST(3,10,0.3,FALSE)", 0.266827932),
        ("BINOM.DIST(3,10,0.3,TRUE)", 0.6496107184),
        ("BINOM.DIST(50,100,0.5,FALSE)", 0.07958923738717878),
        ("BINOM.DIST(0,1000,0.001,TRUE)", 0.36769542477096406),
        ("BINOM.DIST(0,10,0,FALSE)", 1.0),
        ("BINOM.DIST.RANGE(60,0.75,48)", 0.0839749674290475),
        ("BINOM.DIST.RANGE(60,0.75,45,50)", 0.5236297934718872),
        ("BINOM.DIST.RANGE(10,0.3,0,10)", 1.0),
        ("BINOM.INV(6,0.5,0.75)", 4.0),
        ("CRITBINOM(6,0.5,0.75)", 4.0),
        ("BINOM.INV(1000,0.01,0.999)", 21.0),
        ("BINOM.INV(0,0.3,0.5)", 0.0),
    ]);
    check(&[
        ("BINOM.DIST(11,10,0.3,FALSE)", e(ErrorValue::Num)),
        ("BINOM.DIST(3,10,1.2,FALSE)", e(ErrorValue::Num)),
        ("BINOM.DIST(3,10,0.3)", e(ErrorValue::Value)),
        ("BINOM.DIST.RANGE(10,0.3,5,4)", e(ErrorValue::Num)),
        ("BINOM.INV(10,0.3,0)", e(ErrorValue::Num)),
        ("BINOM.INV(10,0.3,1)", e(ErrorValue::Num)),
        ("BINOM.INV(10,0,0.5)", e(ErrorValue::Num)),
        ("BINOM.INV(10,1,0.5)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn poisson_hypergeometric_and_negative_binomial() {
    approx(&[
        ("POISSON.DIST(2,5,FALSE)", 0.08422433748856833),
        ("POISSON.DIST(2,5,TRUE)", 0.12465201948308113),
        ("POISSON(2,5,TRUE)", 0.12465201948308113),
        ("POISSON.DIST(100,100,TRUE)", 0.5265621985299984),
        ("POISSON.DIST(0,0.5,FALSE)", 0.6065306597126334),
        ("POISSON.DIST(3,0,TRUE)", 1.0),
        ("HYPGEOM.DIST(1,4,8,20,TRUE)", 0.46542827657378727),
        ("HYPGEOM.DIST(1,4,8,20,FALSE)", 0.3632610939112486),
        ("HYPGEOMDIST(1,4,8,20)", 0.3632610939112486),
        ("HYPGEOM.DIST(3,15,8,20,TRUE)", 0.003611971104231163),
        ("HYPGEOM.DIST(4,4,8,20,TRUE)", 1.0),
        ("NEGBINOM.DIST(10,5,0.25,FALSE)", 0.05504866037517785),
        ("NEGBINOM.DIST(10,5,0.25,TRUE)", 0.31351405847817665),
        ("NEGBINOMDIST(10,5,0.25)", 0.05504866037517785),
    ]);
    // outside its support the hypergeometric answers zero, not an error
    check(&[
        ("HYPGEOM.DIST(5,4,8,20,FALSE)", n(0.0)),
        ("HYPGEOM.DIST(5,4,8,20,TRUE)", n(1.0)),
        ("HYPGEOM.DIST(0,15,8,20,FALSE)", n(0.0)),
        ("HYPGEOM.DIST(0,15,8,20,TRUE)", n(0.0)),
        ("HYPGEOM.DIST(1,4,0,20,FALSE)", n(0.0)),
        ("HYPGEOM.DIST(-1,4,8,20,FALSE)", e(ErrorValue::Num)),
        ("HYPGEOM.DIST(1,21,8,20,FALSE)", e(ErrorValue::Num)),
        ("POISSON.DIST(-1,5,TRUE)", e(ErrorValue::Num)),
        ("NEGBINOM.DIST(0,1,1,FALSE)", e(ErrorValue::Num)),
        ("NEGBINOM.DIST(0,1,0,FALSE)", e(ErrorValue::Num)),
        ("NEGBINOM.DIST(2,0,0.5,TRUE)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn continuous_helpers() {
    approx(&[
        ("EXPON.DIST(0.2,10,TRUE)", 0.8646647167633873),
        ("EXPON.DIST(0.2,10,FALSE)", 1.353352832366127),
        ("EXPONDIST(0.2,10,TRUE)", 0.8646647167633873),
        ("EXPON.DIST(1E-10,1,TRUE)", 9.999999999500001e-11),
        ("WEIBULL.DIST(105,20,100,TRUE)", 0.929581390069277),
        ("WEIBULL.DIST(105,20,100,FALSE)", 0.035588864024503564),
        ("WEIBULL(105,20,100,TRUE)", 0.929581390069277),
        ("FISHER(0.75)", 0.9729550745276566),
        ("FISHERINV(0.972955)", 0.7499999673941484),
        ("STANDARDIZE(42,40,1.5)", 1.3333333333333333),
        ("Z.TEST({3,6,7,8,6,5,4,2,1,9},4)", 0.09057419685136381),
        ("Z.TEST({3,6,7,8,6,5,4,2,1,9},6)", 0.86304338912953),
        ("ZTEST({3,6,7,8,6,5,4,2,1,9},4,2)", 0.04099516050019153),
        ("PERMUT(100,3)", 970_200.0),
        ("PERMUT(0,0)", 1.0),
        ("PERMUTATIONA(3,2)", 9.0),
        ("PERMUTATIONA(0,0)", 1.0),
    ]);
    check(&[
        ("FISHER(1)", e(ErrorValue::Num)),
        ("STANDARDIZE(42,40,0)", e(ErrorValue::Num)),
        ("PERMUT(2,3)", e(ErrorValue::Num)),
        ("PERMUT(-1,0)", e(ErrorValue::Num)),
    ]);
}

#[test]
fn ranks_percentiles_and_spread() {
    approx(&[
        ("PERCENTILE.EXC({1,2,3,4,5,6,7,8,9},0.25)", 2.5),
        ("QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},1)", 15.0),
        ("PERCENTRANK.INC({1,2,3,4,5,6,7,8,9,10},4)", 0.333),
        ("PERCENTRANK({1,2,3,4,5,6,7,8,9,10},20/3,2)", 0.62),
        ("PERCENTRANK.INC({1,1,3},2)", 0.75),
        ("PERCENTRANK.EXC({1,1,3},2)", 0.625),
        ("PERCENTRANK.EXC({1,2,3,4,5,6,7,8,9,10},8)", 0.727),
        ("PERCENTRANK.INC({5},5)", 1.0),
        ("SKEW.P({3,4,5,2,3,4,5,6,4,7})", 0.3031933393541438),
        ("VARA({1,2,3})", 1.0),
        ("VARPA({1,2,3})", 0.6666666666666666),
        ("STDEVPA({1,2,3})", 0.816496580927726),
        ("STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 3.305718950210041),
        ("PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)", 0.1),
        ("PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)", 0.8),
        ("PROB({0,1},{1.5,-0.5},0)", 1.5),
        ("AGGREGATE(18,6,{1,2,3,4,5,6,7,8,9},0.25)", 2.5),
        ("AGGREGATE(19,6,{6,7,15,36,39,40,41,42,43,47,49},1)", 15.0),
    ]);
    check(&[
        (
            "PERCENTILE.EXC({1,2,3,4,5,6,7,8,9},0.05)",
            e(ErrorValue::Num),
        ),
        (
            "QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},0)",
            e(ErrorValue::Num),
        ),
        ("PERCENTRANK.INC({1,2,3},4)", e(ErrorValue::NA)),
        ("PERCENTRANK.INC({1,2,3},2,0)", e(ErrorValue::Num)),
        ("PROB({0,1,2,3},{0.2,0.3,0.1,0.3},1,3)", e(ErrorValue::Num)),
        ("PROB({0,1,2},{0.2,0.3,0.5,0.1},1,3)", e(ErrorValue::NA)),
        ("AGGREGATE(19,6,{1,2,3},4)", e(ErrorValue::Num)),
        ("STEYX({1,2},{3,4})", e(ErrorValue::Div0)),
    ]);
    assert_eq!(
        eval("PEARSON({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})"),
        eval("CORREL({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})")
    );
}

#[test]
fn mode_mult_lists_every_mode_in_order_of_appearance() {
    assert_eq!(
        block("MODE.MULT({1,2,2,3,3,1,4})"),
        (3, 1, vec![n(1.0), n(2.0), n(3.0)])
    );
    assert_eq!(block("MODE.MULT({1,2,3})"), (1, 1, vec![e(ErrorValue::NA)]));
}

/// LOGEST fits ln(y); y = 2 * 3^x comes back exactly as m = 3 and b = 2, and
/// GROWTH follows the same curve.
#[test]
fn logest_and_growth_fit_an_exponential() {
    let (rows, cols, cells) = block("LOGEST({6;18;54;162},{1;2;3;4})");
    assert_eq!((rows, cols), (1, 2));
    close(&cells[0], 3.0, "m");
    close(&cells[1], 2.0, "b");
    let (rows, cols, cells) = block("GROWTH({6;18;54;162},{1;2;3;4},{5;6})");
    assert_eq!((rows, cols), (2, 1));
    close(&cells[0], 486.0, "GROWTH at 5");
    close(&cells[1], 1458.0, "GROWTH at 6");
    let (_, _, cells) = block("LOGEST({6;18;54;162},{1;2;3;4},FALSE)");
    assert_eq!(cells[1], n(1.0));
    assert_eq!(block("LOGEST({1;0;3},{1;2;3})").2, vec![e(ErrorValue::Num)]);
    assert_eq!(block("GROWTH({1;-2;3})").2, vec![e(ErrorValue::Num)]);
}

/// a predictor the later ones already span is dropped, as Excel drops it:
/// with x2 = 2 x1 the fit keeps x2, reports 0 for x1, and spends one degree
/// of freedom fewer.
#[test]
fn linest_drops_a_collinear_predictor() {
    let ys = "{5.1;6.9;9.2;10.8;13.1;15;16.8;19.3}";
    let xs = "{1,2;2,4;3,6;4,8;5,10;6,12;7,14;8,16}";
    let (rows, cols, cells) = block(&format!("LINEST({ys},{xs},TRUE,TRUE)"));
    assert_eq!((rows, cols), (5, 3));
    close(&cells[0], 1.0035714285714288, "m2");
    assert_eq!(cells[1], n(0.0));
    close(&cells[2], 2.992857142857142, "b");
    close(&cells[3], 0.015198982541432942, "se m2");
    assert_eq!(cells[4], n(0.0));
    close(&cells[10], 6.0, "df");
}
