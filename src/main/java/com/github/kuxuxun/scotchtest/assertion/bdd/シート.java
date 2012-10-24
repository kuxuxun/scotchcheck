package com.github.kuxuxun.scotchtest.assertion.bdd;

import static org.junit.Assert.assertEquals;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import com.github.kuxuxun.commonutil.ut.TestHelper;
import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.cell.ScColor;
import com.github.kuxuxun.scotchtest.assertion.AssertBorders;
import com.github.kuxuxun.scotchtest.assertion.AssertFont;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertSheetSetting;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertStyle;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertValue;
import com.github.kuxuxun.scotchtest.assertion.assertion.At;
import com.github.kuxuxun.scotchtest.assertion.assertion.In;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.DataFormatExample;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.FontExample;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.CellBorders.Border;

public class シート extends TestHelper {

	private final ScSheet targetReport;
	private In range;

	private List<String> poses = null;

	private Border leftOrAll, top, right, bottom;

	public シート(ScSheet report) {
		this.targetReport = report;
	}

	public ページ設定 のページ設定() {
		return new ページ設定(targetReport, range);
	}

	public シート のセル(String pos) {
		this.range = At.at(pos);
		return this;
	}

	public シート また() {
		return this;
	}

	public シート かつ() {
		return this;
	}

	public シート から(String pos) {
		this.range = new In(range.getPositions().get(0).toReference(), pos);
		return this;
	}

	/**
	 * @param pos
	 * @deprecated 未実装
	 * @return
	 */
	@Deprecated
	public シート と(String pos) {
		if (poses == null) {
			poses = new ArrayList<String>();
			if (range != null) {
				for (ScPos each : range.getPositions()) {
					poses.add(each.toReference());
				}
			}
		}
		poses.add(pos);

		// String[] arr = poses.toArray(a);

		// this.range = Positions.in(arr);
		// TODO 実装
		throw new IllegalAccessError("未実装です");
		// return this;
	}

	public スタイル のスタイル() throws FileNotFoundException, IOException {
		return new スタイル(targetReport, range);
	}

	public シート のそれぞれのセルが罫線で囲まれている() throws Exception {
		AssertBorders.that(targetReport).eachCellsSurroundByThinBorder(range);
		return this;
	}

	public シート のそれぞれのセルが次の罫線で囲まれている(Border b) throws Exception {
		AssertBorders.that(targetReport).eachCellsSurroundBy(b, range);

		return this;
	}

	public void の範囲が細い線で囲まれており中線が無い() throws FileNotFoundException, IOException {
		AssertBorders.that(targetReport).isSurroundByBoxOfThinBorder(range);

	}

	public シート の範囲が線(Border left) {
		return の範囲が線(left, null, null, null);
	}

	public シート の範囲が線(Border left, Border top, Border right, Border bottom) {

		this.leftOrAll = left;
		this.top = top;
		this.right = right;
		this.bottom = bottom;

		return this;
	}

	public シート で囲まれており中線が無い() throws FileNotFoundException, IOException {

		if (top == null) {
			AssertBorders.that(targetReport).isSurroundByBorderWith(leftOrAll,
					range);
		} else {
			AssertBorders.that(targetReport).isSurroundByBorderWith(leftOrAll,
					top, right, bottom, range);
		}
		return this;
	}

	public シート の範囲に中線が無い() throws FileNotFoundException, IOException {
		AssertBorders.that(targetReport).noInternalBorderIn(range);
		return this;
	}

	public シート の左の罫線は(Border b) throws FileNotFoundException, IOException {
		AssertBorders.that(targetReport).isLeftBorder(b, range);
		return this;
	}

	public シート の左に罫線はない() throws FileNotFoundException, IOException {
		の左の罫線は(Border.empty());
		return this;
	}

	public シート の上の罫線は(Border b) throws FileNotFoundException, IOException {
		AssertBorders.that(targetReport).isTopBorder(b, range);
		return this;
	}

	public シート の上に罫線はない() throws FileNotFoundException, IOException {
		の上の罫線は(Border.empty());
		return this;
	}

	public シート の名前は(String expected) throws FileNotFoundException, IOException {

		assertEquals("シート名が", expected, targetReport.getName());
		return this;
	}

	public シート の右の罫線は(Border b) throws FileNotFoundException, IOException {
		AssertBorders.that(targetReport).isRightBorder(b, range);
		return this;
	}

	public シート の右に罫線はない() throws FileNotFoundException, IOException {
		の右の罫線は(Border.empty());
		return this;
	}

	public シート の下の罫線は(Border b) throws FileNotFoundException, IOException {
		AssertBorders.that(targetReport).isBottomBorder(b, range);
		return this;
	}

	public シート の下に罫線はない() throws FileNotFoundException, IOException {
		の下の罫線は(Border.empty());
		return this;
	}

	public シート で囲まれている_中線があるかないかは気にしない() throws FileNotFoundException,
			IOException {

		if (top == null) {
			AssertBorders.that(targetReport).isSurroundByBorderWith(leftOrAll,
					range);
		} else {
			AssertBorders.that(targetReport).isSurroundByBorderWith(leftOrAll,
					top, right, bottom, range);
		}

		return this;

	}

	public シート のフォントは(String fontName) throws FileNotFoundException,
			IOException {

		AssertFont.that(targetReport).fontNameIs(fontName, range);
		return this;
	}

	public シート のフォントサイズは(int size) throws FileNotFoundException, IOException {
		AssertFont.that(targetReport).fontSizeIs(size, range);
		return this;
	}

	public シート のフォントの色は(ScColor color) throws FileNotFoundException,
			IOException {

		AssertFont.that(targetReport).fontColorIs(color, range);
		return this;

	}

	public シート のフォントはBold() throws FileNotFoundException, IOException {
		AssertFont.that(targetReport).fontIsBold(range);
		return this;

	}

	public シート のフォントはItalic() throws FileNotFoundException, IOException {
		AssertFont.that(targetReport).fontIsItalic(range);
		return this;

	}

	public シート の文字に一本下線がついている() throws FileNotFoundException, IOException {

		AssertFont.that(targetReport).fontIsUnderlinedWithSingle(range);

		return this;

	}

	public シート の文字に二本下線がついている() throws FileNotFoundException, IOException {

		AssertFont.that(targetReport).fontIsUnderlinedWithDouble(range);

		return this;

	}

	public シート フォントは(FontExample font) throws FileNotFoundException,
			IOException {
		return のフォントは(font);
	}

	public シート のフォントは(FontExample font) throws FileNotFoundException,
			IOException {
		AssertFont.that(targetReport).isSameAs(font, range);

		return this;

	}

	public シート の背景色は(ScColor color) throws FileNotFoundException, IOException {
		AssertStyle.that(targetReport).filledColorIs(color, range);
		return this;

	}

	public シート の背景色は設定されていないか自動になっている() throws FileNotFoundException,
			IOException {
		AssertStyle.that(targetReport)
				.filledColorIs(ScColor.NON_OR_AUTO, range);

		return this;

	}

	public シート フォーマットが文字列() throws FileNotFoundException, IOException {
		return のフォーマットが文字列();
	}

	public シート のフォーマットが文字列() throws FileNotFoundException, IOException {
		AssertStyle.that(targetReport).dataFormatIs(
				DataFormatExample.TEXT.getFormatString(), range);
		return this;
	}

	public シート フォーマットが標準() throws FileNotFoundException, IOException {
		return のフォーマットが標準();
	}

	public シート のフォーマットが標準() throws FileNotFoundException, IOException {
		AssertStyle.that(targetReport).dataFormatIs(
				DataFormatExample.GENERAL.getFormatString(), range);
		return this;

	}

	public シート のフォーマットがカンマなし数値() throws FileNotFoundException, IOException {
		AssertStyle.that(targetReport).dataFormatIs(
				DataFormatExample.INTEGER_NEGRED.getFormatString(), range);
		return this;

	}

	public シート のフォーマットが数値_カンマ_負数は赤() throws FileNotFoundException, IOException {
		AssertStyle.that(targetReport)
				.dataFormatIs(
						DataFormatExample.INTEGER_COMMA_NEGRED
								.getFormatString(), range);
		return this;

	}

	public シート のフォーマットが数値_カンマ_負数は赤で最少桁が0() throws FileNotFoundException,
			IOException {
		AssertStyle.that(targetReport)
				.dataFormatIs(
						DataFormatExample.INTEGER_COMMA_NEGRED_LAST_0
								.getFormatString(), range);
		return this;

	}

	public シート のフォーマットが小数点以下2桁_負数は赤() throws FileNotFoundException, IOException {
		AssertStyle.that(targetReport).dataFormatIs(

		DataFormatExample.INTEGER_AFTER_POINT_TWO_NEGRED.getFormatString(),
				range);

		return this;

	}

	public シート のフォーマットが(String formatString) throws FileNotFoundException,
			IOException {
		AssertStyle.that(targetReport).dataFormatIs(formatString, range);

		return this;

	}

	public シート のフォーマットが1000単位で表示_カンマ_負数は赤() throws FileNotFoundException,
			IOException {
		AssertStyle.that(targetReport).dataFormatIs(

		DataFormatExample.INTEGER_IN_KILO_NEGRED.getFormatString(), range);

		return this;

	}

	public シート のフォーマットが少数点2桁を四捨五入したパーセントカンマ表示で表示で負数は赤()
			throws FileNotFoundException, IOException {
		AssertStyle.that(targetReport).dataFormatIs(
				DataFormatExample.PERCENT_ROUND_AT_SEC_NEGRED.getFormatString()

				, range);

		return this;

	}

	public シート の文字列が(String text) throws Exception {
		AssertValue.that(targetReport).hasText("", text, range);
		return this;
	}

	public シート の日付が(String yyyymmdd) throws Exception {
		AssertValue.that(targetReport).hasDate(
				new SimpleDateFormat("yyyyMMdd")
						.parse(removeDateSeparator(yyyymmdd)), range, "");
		return this;
	}

	private static String removeDateSeparator(String s) {
		return s.replaceAll("/", "").replaceAll("-", "");
	}

	public シート の小数点２桁以上の数値が(BigDecimal dec) throws Exception {
		の数値が(dec, 2);
		return this;
	}

	public シート の整数値が(int num) throws Exception {
		AssertValue.that(targetReport).hasNumber("整数値の値 が正しいこと",
				new BigDecimal(num), 0, range);

		return this;
	}

	public シート の式が(String expected) throws Exception {
		AssertValue.that(targetReport).hasFormula("式が正しい事", expected, range);
		return this;
	}

	public シート の数値が(BigDecimal dec, int figAferDecimal) throws Exception {
		AssertValue.that(targetReport).hasNumber("数値の値 が正しいこと", dec,
				figAferDecimal, range);

		return this;
	}

	public シート の値が空() throws Exception {
		AssertValue.that(targetReport).isEmpty("", range);
		return this;
	}

	public 範囲チェック の範囲が() throws Exception {
		return new 範囲チェック(targetReport, range);
	}

	public void が保護されている() {
		AssertSheetSetting.that(targetReport).isProtected();
	}

	public void は保護されていない() {
		AssertSheetSetting.that(targetReport).isProtected();
	}

	public void のセルがロックされている() throws Exception {
		AssertStyle.that(targetReport).isLocked(true, range);
	}

	public void のセルがロックされていない() throws Exception {
		AssertStyle.that(targetReport).isLocked(false, range);
	}
}
