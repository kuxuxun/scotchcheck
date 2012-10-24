package com.github.kuxuxun.scotchtest.assertion;

import org.junit.Before;
import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestExpected;
import com.github.kuxuxun.commonutil.ut.TestHelper;
import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertValue;
import com.github.kuxuxun.scotchtest.assertion.assertion.At;
import com.github.kuxuxun.scotchtest.assertion.assertion.In;
import com.github.kuxuxun.scotchtest.assertion.assertion.Positions;
import com.github.kuxuxun.scotchtest.assertion.bdd.シート;

public class AssertValueTest extends TestHelper {

	static ScSheet report = TestExpected.getFirstSheetOf("expected.xls");

	@Before
	public void setup() {
	}

	@Test
	public void assertRangeValue() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("F18")
				.から("F21").の文字列が("テスト");

		AssertValue.that(report).hasText("範囲内の文字が正しい", "テスト",
				new In("F18", "F21"));
	}

	@Test
	public void assertPositionsValue() throws Exception {
		AssertValue.that(report).hasText("範囲内の文字が正しい", "テスト",
				Positions.in(new String[] { "F18", "F23" }));
	}

	@Test
	public void assertFixCellValue() throws Exception {
		AssertValue.that(report)
				.hasText("正しい文字列の値", "MSPゴシック-10", At.at("C10"));
	}

	@Test(expected = IllegalArgumentException.class)
	public void assertNotTextCellValue() throws Exception {
		AssertValue.that(report).hasText("日付が入っているセルに文字列を期待するとエラー",
				"2011/11/11", At.at("E10"));
	}

	@Test
	public void assertDateCellValue() throws Exception {
		AssertValue.that(report).hasDate(date("2011/11/11"), At.at("E10"),
				"日付の値 が正しいこと");
	}

	@Test
	public void assertNumericCellValue() throws Exception {
		AssertValue.that(report).hasNumber("数値の値 が正しいこと", decimal("-1111"), 2,
				At.at("E11"));
	}

	@Test
	public void assertFormulaCellValue() throws Exception {
		AssertValue.that(report).hasFormula("式の値 が正しいこと", "SUM(E11:E12)",
				At.at("E13"));
	}

}
