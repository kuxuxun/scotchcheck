package com.github.kuxuxun.scotchtest.assertion.assertion;

import static org.junit.Assert.assertEquals;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.cell.ScCell;
import com.github.kuxuxun.scotch.excel.cell.ScColor;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.HAlign;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.VAlign;

public class AssertStyle {
	private static String n2b(String s) {
		if (s == null)
			return "";

		return s;
	}

	private final ScSheet sheet;

	private AssertStyle(ScSheet sheet) {
		this.sheet = sheet;
	}

	public static AssertStyle that(ScSheet sheet) {
		return new AssertStyle(sheet);
	}

	public String makeDesc(String description) {
		String desc = "";
		if (n2b(description).length() != 0) {
			desc = description + "として";
		}

		return desc;
	}

	public void isLocked(final boolean expected, In range)
			throws FileNotFoundException, IOException {
		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "のセルのロック: ",
						expected, cell.getStyle().getStyleWithoutFont()
								.isLocked());
			}
		}.doAssert();

	}

	public void filledColorIs(final ScColor expected, In range)
			throws FileNotFoundException, IOException {
		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "の背景色: ", expected,
						cell.getStyle().getStyleWithoutFont().getFgColor());
			}
		}.doAssert();

	}

	public void dataFormatIs(final String expected, In range)
			throws FileNotFoundException, IOException {

		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "の形式: ", expected,
						cell.getStyle().getStyleWithoutFont()
								.getDataFormatString());
			}
		}.doAssert();

	}

	public void horizontalAlignIs(final HAlign expected, In range)
			throws FileNotFoundException, IOException {

		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "の横Align: ",
						expected, HAlign.getFromCell(cell));
			}
		}.doAssert();

	}

	public void verticalAllignIs(final VAlign expected, In range)
			throws FileNotFoundException, IOException {

		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "の縦Align: ",
						expected, VAlign.getFromCell(cell));
			}
		}.doAssert();

	}

	public void colWidthInSettingVal(final BigDecimal expectedWidth, In range)
			throws FileNotFoundException, Exception {

		// 72ポイント = 1インチ(フォントの高さ) ,1/256 * 文字の大きさ => poiの幅
		// 8.38ポイント => 数字が8.38個入る長さ
		// 11 / 72 * 25.4 => 3.88055 mm
		// 11(font(高さ) 2048(poiの幅) => 8.38 (標準スタイルの"0"が8.38個表示できる)
		// 8.38 => 72ピクセル

		// 8.38p => 72px => 2048
		// 8.88p => 76px => 2432
		// 10p => 85px => 2720
		// 0.5p => 382 ,1p => 768
		// 1p => 13px => 416(ｺﾞｼｯｸ)
		// 1p => 18px(丸ｺﾞｼｯｸ) 418
		// 1p => 8.5px
		// 1文字幅 8 18px => 418
		// 1文字幅 5 => 302

		// 30pix 6 pix => 698

		// setting val = 32 * (n * 8 + 5)

		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "の列幅: ",
						expectedWidth, cell.getStyle().getColWidth());
			}
		}.doAssert();

	}
	// FIXME 列幅
	// FIXME 標準スタイルの確認(列幅算定に使われるため)
}
