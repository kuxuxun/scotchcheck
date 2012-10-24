package com.github.kuxuxun.scotchtest.assertion.bdd;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;

import com.github.kuxuxun.commonutil.ut.TestHelper;
import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertPageSetting;
import com.github.kuxuxun.scotchtest.assertion.assertion.In;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.PageFitting;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.PageNumberPrint;

public class ページ設定 extends TestHelper {

	private final ScSheet targetSheet;

	public ページ設定(ScSheet targetSheet, In range) {
		this.targetSheet = targetSheet;
	}

	public ページ設定 の左マージンは(double d) throws Exception {
		AssertPageSetting.that(targetSheet).leftMarginIs(
				decimal(Double.toString(d)));
		return this;

	}

	public ページ設定 の上マージンは(double d) throws Exception {
		AssertPageSetting.that(targetSheet).topMarginIs(
				decimal(Double.toString(d)));
		return this;
	}

	public ページ設定 の下マージンは(double d) throws Exception {
		AssertPageSetting.that(targetSheet).bottomMarginIs(
				decimal(Double.toString(d)));
		return this;

	}

	public ページ設定 の右マージンは(double d) throws Exception {
		AssertPageSetting.that(targetSheet).rightMarginIs(
				decimal(Double.toString(d)));
		return this;

	}

	public ページ設定 のヘッダマージンは(double d) throws Exception {
		AssertPageSetting.that(targetSheet).headerMarginIs(
				decimal(Double.toString(d)));
		return this;

	}

	public ページ設定 のフッタマージンは(double d) throws Exception {
		AssertPageSetting.that(targetSheet).footerMarginIs(
				decimal(Double.toString(d)));
		return this;

	}

	public ページ設定 の用紙サイズは(String paperSize) throws Exception {
		AssertPageSetting.that(targetSheet).pagePaperSizeIs(paperSize);
		return this;
	}

	public ページ設定 の用紙の向きは(String ori) throws Exception {
		AssertPageSetting.that(targetSheet).pageOrientation(ori);
		return this;
	}

	public ページ設定 は次の拡大率で印刷される(BigDecimal dec) throws Exception {
		AssertPageSetting.that(targetSheet).printWithScaling(dec);

		// hssfとは違いxssf ではチェックがついてなくても取得される模様
		// AssertPageSetting.that(targetSheet).printPageFitTo(
		// PageFitting.pageNotFit());

		return this;

	}

	public ページ設定 は次の縦横ページに合わせて印刷される(int v, int h) throws Exception {
		AssertPageSetting.that(targetSheet).printPageFitTo(
				PageFitting.pageFitToBoth(v, h));

		return this;

	}

	public ページ設定 は次の縦ページにのみ合わせて印刷される(int v) throws Exception {
		AssertPageSetting.that(targetSheet).printPageFitTo(
				PageFitting.pageFitToVert(v));

		return this;

	}

	public ページ設定 は次の横ページにのみ合わせて印刷される(int h) throws Exception {
		AssertPageSetting.that(targetSheet).printPageFitTo(
				PageFitting.pageFitToHori(h));

		return this;

	}

	public ページ設定 のフッターに次の形式でページ番号が表示されている(String infix, String prefix)
			throws FileNotFoundException, IOException {
		AssertPageSetting.that(targetSheet).printPageNumberSettingIs(
				PageNumberPrint.pageOfAllAtCenterFooter(infix, prefix));
		return this;
	}

}
