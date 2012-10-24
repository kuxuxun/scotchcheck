package com.github.kuxuxun.scotchtest.assertion;

import static org.junit.Assert.assertTrue;

import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestExpected;
import com.github.kuxuxun.scotch.excel.cell.ScColor;
import com.github.kuxuxun.scotchtest.assertion.assertion.In;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.CellBorders;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.CellBorders.Border;
import com.github.kuxuxun.scotchtest.assertion.bdd.シート;

public class AssertBorderTest {

	@Test
	public void 結合されている１行分セルのボーダーの指定が間違っていた場合正しくエラーとなる() throws Exception {

		boolean exeptionOccured = false;
		try {

			// 罫線はB22からD22範囲についている
			new シート(TestExpected.getFirstSheetOf("compareTarget.xls")).のセル(
					"B22").から("E22").の範囲が線(Border.medium(ScColor.BLACK))
					.で囲まれており中線が無い();

		} catch (Throwable e) {
			exeptionOccured = true;
		}

		assertTrue(exeptionOccured);

	}

	@Test
	public void AllEachSurroundByborderTest() throws Exception {
		new シート(TestExpected.getFirstSheetOf("compareTarget.xls")).のセル("H6")
				.から("I8").のそれぞれのセルが罫線で囲まれている();

		AssertBorders.that(TestExpected.getFirstSheetOf("compareTarget.xls"))
				.eachCellsSurroundByThinBorder(new In("H6", "I8"));

	}

	@Test
	public void surroundByborderTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("compareTarget.xls")).のセル("H12")
				.から("I15").の範囲が細い線で囲まれており中線が無い();

		AssertBorders.that(TestExpected.getFirstSheetOf("compareTarget.xls"))
				.isSurroundByBoxOfThinBorder(new In("H12", "I15"));
	}

	@Test
	public void surroundByborderLeftEdgeTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("A17").から(
				"B18").の範囲が細い線で囲まれており中線が無い();

		AssertBorders.that(TestExpected.getFirstSheetOf("expected.xls"))
				.isSurroundByBoxOfThinBorder(new In("A17", "B18"));
	}

	@Test
	public void surroundByborderTopEdgeTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("K1")
				.から("L2").の範囲が細い線で囲まれており中線が無い();

		AssertBorders.that(TestExpected.getFirstSheetOf("expected.xls"))
				.isSurroundByBoxOfThinBorder(new In("K1", "L2"));
	}

	@Test
	public void surroundBySeveralBordersTest() throws Exception {

		// FIXME 線もDSLっぽく
		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("H18").から(
				"I20").の範囲が線(Border.thin(ScColor.RED),
				Border.medium(ScColor.BLACK), Border.thin(ScColor.BLACK),
				Border.medium(ScColor.RED)).で囲まれており中線が無い();

		AssertBorders.that(TestExpected.getFirstSheetOf("expected.xls"))
				.isSurroundByBorderWith(Border.thin(ScColor.RED),
						Border.medium(ScColor.BLACK),
						Border.thin(ScColor.BLACK), Border.medium(ScColor.RED),
						new In("H18", "I20"));

	}

	@Test
	public void bordersOnEachSideTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("H18").から(
				"I20").の左の罫線は(Border.thin(ScColor.RED));
		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("H18").から(
				"I20").の上の罫線は(Border.medium(ScColor.BLACK));
		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("H18").から(
				"I20").の右の罫線は(Border.thin(ScColor.BLACK));
		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("H18").から(
				"I20").の下の罫線は(Border.medium(ScColor.RED));

		AssertBorders.that(TestExpected.getFirstSheetOf("expected.xls"))
				.isLeftBorder(Border.thin(ScColor.RED), new In("H18", "I20"));

		AssertBorders
				.that(TestExpected.getFirstSheetOf("expected.xls"))
				.isTopBorder(Border.medium(ScColor.BLACK), new In("H18", "I20"));

		AssertBorders
				.that(TestExpected.getFirstSheetOf("expected.xls"))
				.isRightBorder(Border.thin(ScColor.BLACK), new In("H18", "I20"));

		AssertBorders.that(TestExpected.getFirstSheetOf("expected.xls"))
				.isBottomBorder(Border.medium(ScColor.RED),
						new In("H18", "I20"));

	}

	@Test
	public void surroundByBoxTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("H12").から(
				"I15").の範囲が線(Border.thick(ScColor.RED)).で囲まれており中線が無い();

		AssertBorders.that(TestExpected.getFirstSheetOf("expected.xls"))
				.isSurroundInBoxWith(CellBorders.Border.thick(ScColor.RED),
						new In("H12", "I15"));
	}

	@Test
	public void surroundByBorderTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("K6")
				.から("L8").の範囲が線(Border.thin(ScColor.BLACK))
				.で囲まれている_中線があるかないかは気にしない();

		AssertBorders.that(TestExpected.getFirstSheetOf("expected.xls"))
				.isSurroundByBorderWith(CellBorders.Border.thin(ScColor.BLACK),
						new In("K6", "L8"));
	}
}
