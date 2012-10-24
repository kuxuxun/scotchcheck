package com.github.kuxuxun.scotchtest.assertion;

import org.junit.Before;
import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestExpected;
import com.github.kuxuxun.commonutil.ut.TestHelper;
import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.cell.ScColor;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertStyle;
import com.github.kuxuxun.scotchtest.assertion.assertion.At;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.DataFormatExample;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.HAlign;
import com.github.kuxuxun.scotchtest.assertion.bdd.シート;

public class AssertStyleTest extends TestHelper {

	static ScSheet sheet = TestExpected.getFirstSheetOf("expected.xls");

	@Before
	public void setup() {
	}

	@Test
	public void columnWidth() throws Exception {
		AssertStyle.that(sheet)
				.colWidthInSettingVal(decimal(2048), At.at("A1"));
	}

	@Test
	public void cellBackGroundColorTest() throws Exception {

		new シート(sheet).のセル("C4").の背景色は(ScColor.RED);
		new シート(sheet).のセル("C2").の背景色は設定されていないか自動になっている();

		AssertStyle.that(sheet).filledColorIs(ScColor.RED, At.at("C4"));
		AssertStyle.that(sheet).filledColorIs(ScColor.NON_OR_AUTO, At.at("C2"));
	}

	@Test
	public void cellDataFormatColorTest() throws Exception {

		new シート(sheet).のセル("E14").のフォーマットが数値_カンマ_負数は赤();

		new シート(sheet).のセル("E15").のフォーマットが1000単位で表示_カンマ_負数は赤();

		new シート(sheet).のセル("E16").のフォーマットが少数点2桁を四捨五入したパーセントカンマ表示で表示で負数は赤();

		AssertStyle.that(sheet).dataFormatIs(
				DataFormatExample.INTEGER_COMMA_NEGRED.getFormatString(),
				At.at("E14"));

		AssertStyle.that(sheet).dataFormatIs(

		DataFormatExample.INTEGER_IN_KILO_NEGRED.getFormatString(),
				At.at("E15"));

		AssertStyle.that(sheet).dataFormatIs(
				DataFormatExample.PERCENT_ROUND_AT_SEC_NEGRED.getFormatString()

				, At.at("E16"));
	}

	@Test
	public void cellAlign() throws Exception {
		AssertStyle.that(sheet).horizontalAlignIs(HAlign.RIGHT, At.at("E16"));
	}

}
