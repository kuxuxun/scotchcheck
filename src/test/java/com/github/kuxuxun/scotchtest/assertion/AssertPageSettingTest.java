package com.github.kuxuxun.scotchtest.assertion;

import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestExpected;
import com.github.kuxuxun.commonutil.ut.TestHelper;
import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.style.ScMargin;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertPageSetting;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.PageFitting;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.PageNumberPrint;
import com.github.kuxuxun.scotchtest.assertion.bdd.シート;

public class AssertPageSettingTest extends TestHelper {

	static ScSheet report = TestExpected.getFirstSheetOf("expected.xls");

	@Test
	public void incheToCentimaterTest() throws Exception {

		assertCompareEquals(decimal(2), ScMargin
				.inchToCentiMeter(decimal(0.75)));

		assertCompareEquals(decimal(2.5), ScMargin
				.inchToCentiMeter(decimal(1.0)));

		assertCompareEquals(decimal(1.3), ScMargin
				.inchToCentiMeter(decimal(0.5118)));

		assertCompareEquals(decimal(1.1), ScMargin
				.inchToCentiMeter(decimal(0.4330708661417323)));

	}

	@Test
	public void headerMarginTest() throws Exception {

		new シート(report).のページ設定().の上マージンは(1.1);
		new シート(report).のページ設定().の左マージンは(2.2);
		new シート(report).のページ設定().の下マージンは(3.3);
		new シート(report).のページ設定().の右マージンは(4.4);
		new シート(report).のページ設定().のヘッダマージンは(5.5);
		new シート(report).のページ設定().のフッタマージンは(6.6);

		AssertPageSetting.that(report).topMarginIs(decimal(1.1));
		AssertPageSetting.that(report).leftMarginIs(decimal(2.2));
		AssertPageSetting.that(report).bottomMarginIs(decimal(3.3));
		AssertPageSetting.that(report).rightMarginIs(decimal(4.4));
		AssertPageSetting.that(report).headerMarginIs(decimal(5.5));
		AssertPageSetting.that(report).footerMarginIs(decimal(6.6));
	}

	@Test
	public void pagePaperSizeTest() throws Exception {

		new シート(report).のページ設定().の用紙サイズは("A4");

		AssertPageSetting.that(report).pagePaperSizeIs("A4");
	}

	@Test
	public void pageOrientationSizeTest() throws Exception {

		new シート(report).のページ設定().の用紙の向きは("縦");

		AssertPageSetting.that(report).pageOrientation("縦");
	}

	@Test
	public void pageFitWhenPrintTest() throws Exception {

		new シート(report).のページ設定().は次の縦横ページに合わせて印刷される(2, 1);

		AssertPageSetting.that(report).printPageFitTo(
				PageFitting.pageFitToBoth(2, 1));
	}

	@Test
	public void footerPrintPageNumTest() throws Exception {

		new シート(report).のページ設定().のフッターに次の形式でページ番号が表示されている(" / ", " ページ");

		AssertPageSetting.that(report).printPageNumberSettingIs(
				PageNumberPrint.pageOfAllAtCenterFooter(" / ", " ページ"));

	}
}
