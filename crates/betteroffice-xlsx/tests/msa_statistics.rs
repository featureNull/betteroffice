//! Every formula of the measurement system analysis fixture, evaluated by the
//! engine on its own and compared with the result Excel cached for it. The
//! fixture is the Bosch Heft 10 (Verfahren 1) template plus sheets of
//! statistics formulas; `scripts/create-xlsx-statistics-fixture.py` builds it
//! and has Excel compute the cached values.

use xlsx_calc::{EvalContext, Value, evaluate, evaluate_array, parse_formula};
use xlsx_model::{CellProvider, CellRef, CellValue, SheetId, Workbook};

const FIXTURE: &[u8] = include_bytes!("fixtures/msa-statistics.xlsx");

/// agreement asked of every number, relative to Excel's value.
const RELATIVE: f64 = 1e-9;

fn workbook() -> Workbook {
    xlsx_parse::parse_workbook(&ooxml_opc::unzip_parts(FIXTURE).unwrap()).unwrap()
}

fn agrees(got: &CellValue, want: &CellValue) -> bool {
    match (got, want) {
        (CellValue::Number { value: got }, CellValue::Number { value: want }) => {
            (got - want).abs() <= RELATIVE * want.abs() || (got - want).abs() <= 1e-15
        }
        _ => got == want,
    }
}

fn sheet_id(wb: &Workbook, name: &str) -> SheetId {
    SheetId(wb.sheets.iter().position(|s| s.name == name).unwrap() as u32)
}

/// the engine's answer for one formula cell, as the block it spills.
fn evaluate_cell(wb: &Workbook, sheet: SheetId, at: CellRef, formula: &str, array: bool) -> Value {
    let expr = parse_formula(formula).unwrap_or_else(|e| panic!("{formula}: {e:?}"));
    let mut ctx = EvalContext::new(wb, sheet);
    ctx.cell = Some(at);
    if array {
        evaluate_array(&expr, &ctx)
    } else {
        Value::Scalar(evaluate(&expr, &ctx))
    }
}

#[test]
fn every_formula_matches_the_result_excel_cached() {
    let wb = workbook();
    let mut checked = 0;
    let mut mismatches = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        let id = SheetId(index as u32);
        for (at, cell) in sheet.iter_cells() {
            let Some(formula) = &cell.formula else {
                continue;
            };
            let spill = sheet.array_formula(at);
            let value = evaluate_cell(&wb, id, at, formula, spill.is_some());
            let block = value.into_array();
            let (rows, cols) = match spill {
                Some(range) => (
                    (range.end.row - range.start.row + 1) as usize,
                    (range.end.col - range.start.col + 1) as usize,
                ),
                None => (1, 1),
            };
            if (block.rows(), block.cols()) != (rows, cols) {
                mismatches.push(format!(
                    "{}!{} {formula}: {}x{} block, Excel spilled {rows}x{cols}",
                    sheet.name,
                    at.to_a1(),
                    block.rows(),
                    block.cols()
                ));
                continue;
            }
            for row in 0..rows {
                for col in 0..cols {
                    let target = CellRef::new(at.row + row as u32, at.col + col as u32);
                    let got = block.at(row, col);
                    let want = wb.value(id, target);
                    checked += 1;
                    if !agrees(&got, &want) {
                        mismatches.push(format!(
                            "{}!{} {formula}: engine {got:?}, Excel {want:?}",
                            sheet.name,
                            target.to_a1()
                        ));
                    }
                }
            }
        }
    }
    assert!(checked > 400, "checked only {checked} values");
    assert!(
        mismatches.is_empty(),
        "{} of {checked} values disagree:\n{}",
        mismatches.len(),
        mismatches.join("\n")
    );
}

/// the figures Heft 10 reports for its Verfahren 1 example.
#[test]
fn the_capability_study_reproduces_the_heft_10_example() {
    let wb = workbook();
    let sheet = sheet_id(&wb, "Berechnung");
    let result = |a1: &str| {
        let at = CellRef::parse_a1(a1).unwrap();
        let formula = wb.sheets[sheet.0 as usize]
            .cell(at)
            .unwrap()
            .formula
            .as_ref()
            .unwrap();
        evaluate_cell(&wb, sheet, at, formula, false).into_scalar()
    };
    let close = |a1: &str, want: f64, tolerance: f64| match result(a1) {
        CellValue::Number { value } => {
            assert!((value - want).abs() <= tolerance, "{a1}: {value} != {want}")
        }
        other => panic!("{a1}: {other:?}"),
    };
    close("B19", 6.0009, 1e-12);
    close("B20", 0.000994884876941735, 1e-15);
    close("B27", 2.01028, 1e-5);
    close("B28", 1.64173, 1e-5);
    close("B29", 0.016667, 1e-6);
    close("B42", 0.000282743153607097, 1e-15);
    assert_eq!(
        result("B43"),
        CellValue::Text {
            value: "signifikant".into()
        }
    );
}
