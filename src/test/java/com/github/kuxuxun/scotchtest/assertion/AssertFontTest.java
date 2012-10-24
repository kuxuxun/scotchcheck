package com.github.kuxuxun.scotchtest.assertion;

import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestExpected;
import com.github.kuxuxun.scotch.excel.cell.ScColor;
import com.github.kuxuxun.scotchtest.assertion.assertion.At;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.FontExample;
import com.github.kuxuxun.scotchtest.assertion.bdd.シート;

public class AssertFontTest {
	@Test
	public void FontNameTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C10")
				.のフォントは("ＭＳ Ｐゴシック");

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C11")
				.のフォントは("Arial");

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls"))
				.fontNameIs("ＭＳ Ｐゴシック", At.at("C10"));

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls"))
				.fontNameIs("Arial", At.at("C11"));
	}

	@Test
	public void FontSizeTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C10")
				.のフォントサイズは(10);

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C11")
				.のフォントサイズは(11);

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls"))
				.fontSizeIs(10, At.at("C10"));

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls"))
				.fontSizeIs(11, At.at("C11"));
	}

	@Test
	public void FontColorTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C11")
				.のフォントの色は(ScColor.RED);

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls"))
				.fontColorIs(ScColor.RED, At.at("C11"));
	}

	@Test
	public void FontBoldTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C11")
				.のフォントはBold();

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls"))
				.fontIsBold(At.at("C11"));

	}

	@Test
	public void fontItalicTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C11")
				.のフォントはItalic();

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls"))
				.fontIsItalic(At.at("C11"));
	}

	@Test
	public void fontSingleUnderilneTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C11")
				.の文字に一本下線がついている();

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls"))
				.fontIsUnderlinedWithSingle(At.at("C11"));
	}

	@Test
	public void fontDoubleUnderilneTest() throws Exception {

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C10")
				.の文字に二本下線がついている();
		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls"))
				.fontIsUnderlinedWithDouble(At.at("C10"));
	}

	@Test
	public void assertPositiveFontByExampleWithACellTest() throws Exception {

		FontExample c10 = FontExample.Builder.simpleMSPGothic(10)
				.setUnderLineDouble(true).setColor(ScColor.BLACK).build();

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C10")
				.のフォントは(c10);

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls")).isSameAs(
				c10, At.at("C10"));

		FontExample c11 = FontExample.Builder.simpleArial(11).setBold(true)
				.setUnderLineSingle(true).setColor(ScColor.RED).build();

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C11")
				.のフォントは(c11);

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls")).isSameAs(
				c11, At.at("C11"));

	}

	@Test
	public void assertNegativeFontByExampleWithACellTest() throws Exception {
		FontExample negative = FontExample.Builder.simpleMSPGothic(10)
				.setUnderLineSingle(false).setBold(false).build();

		new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C10")
				.のフォントは(negative);

		AssertFont.that(TestExpected.getFirstSheetOf("expected.xls")).isSameAs(
				negative, At.at("C10"));

	}

}
