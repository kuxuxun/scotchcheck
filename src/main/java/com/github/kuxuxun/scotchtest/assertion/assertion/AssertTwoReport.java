package com.github.kuxuxun.scotchtest.assertion.assertion;

import static org.junit.Assert.assertEquals;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import com.github.kuxuxun.commonutil.ut.TestExpected;
import com.github.kuxuxun.commonutil.ut.TestHelper;
import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.ScWorkbook;
import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.cell.ScCell;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.BordersInSheet;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.CellBorders;

/**
 * 二つのexcel帳票を比較し、差異の比較を行います。 差異がある場合はその旨を表示します。
 * 
 * @author
 * @deprecated
 */
@Deprecated
public class AssertTwoReport extends TestHelper {

	private String expectedReportFileName;
	private final ScSheet targetSheet;

	private boolean onlyValueAndBorder = false;

	public AssertTwoReport(ScSheet sheet) {
		this("", sheet);
	}

	public AssertTwoReport(String brief, ScSheet sheet) {
		this.targetSheet = sheet;
	}

	public AssertTwoReport isEqualsInValueAndStylesTo(String exptectedFile) {
		this.expectedReportFileName = exptectedFile;
		return this;
	}

	public AssertTwoReport isEqualsInValueAndBorderTo(String exptectedFile) {
		this.expectedReportFileName = exptectedFile;
		this.onlyValueAndBorder = true;
		return this;
	}

	private class Comparator {

		private final ScSheet expectedSheet;
		private final ScSheet compareTargetSheet;

		public Comparator(ScSheet expectedReport, ScSheet compareTargetReport)
				throws FileNotFoundException, IOException {

			this.expectedSheet = expectedReport;
			this.compareTargetSheet = compareTargetReport;
		}

		public void compareBorderLookAndFeel(String topLeft, String bottomRight) {

			BordersInSheet expectedBorders = BordersInSheet.fetchFromSheet(
					topLeft, bottomRight, this.expectedSheet);
			BordersInSheet compareBorders = BordersInSheet.fetchFromSheet(
					topLeft, bottomRight, this.compareTargetSheet);

			assertEquals(new ArrayList<CellBorders>(), expectedBorders
					.getLookAndFeelDifferenceWith(compareBorders));

		}

		public void compareOnlyValueIn(String topLeft, String bottomRight) {

			CellReference topLeftCell = new CellReference(topLeft);
			CellReference bottomRightCell = new CellReference(bottomRight);

			for (int row = topLeftCell.getRow(); row <= bottomRightCell
					.getRow(); row++) {
				for (int col = topLeftCell.getCol(); col <= bottomRightCell
						.getCol(); col++) {

					compareCellValueEachOther(new ScPos(row, col));
				}
			}

			compareBorderLookAndFeel(topLeft, bottomRight);
		}

		public void comparaValueAndStyleIn(String topLeft, String bottomRight) {

			CellReference topLeftCell = new CellReference(topLeft);
			CellReference bottomRightCell = new CellReference(bottomRight);

			for (int row = topLeftCell.getRow(); row <= bottomRightCell
					.getRow(); row++) {

				// TODO 微妙に違うときある
				// compareRowHeight(row);

				for (int col = topLeftCell.getCol(); col <= bottomRightCell
						.getCol(); col++) {

					// TODO 微妙に違うときある
					// comareColWidth(col);

					compareValueAndStyleEachOther(new ScPos(row, col));
				}
			}

			comparePrintSetting();
			// compareStyleNumberInfo(); fontの数が違って免土井

		}

		private void compareStyleNumberInfo() {

			assertEquals("fontの数: ", expectedSheet.getWorkBook()
					.getPoiWorkBook().getNumberOfFonts(), compareTargetSheet
					.getWorkBook().getPoiWorkBook().getNumberOfFonts());

			assertEquals("styleの数: ", expectedSheet.getWorkBook()
					.getPoiWorkBook().getNumCellStyles(), compareTargetSheet
					.getWorkBook().getPoiWorkBook().getNumCellStyles());

			// TODO
			// assertEquals("dataformatの数: ", expectedSheet.getWorkBook()
			// .getPoiWorkBook()
			// .createDataFormat()
			// .getNumberOfBuiltinBuiltinFormats(), compareTargetSheet
			// .getWorkBook().getPoiWorkBook().createDataFormat()
			// .getNumberOfBuiltinBuiltinFormats());

			// 出るとうざいのでcommment out見たいならコメントをはずす
			// System.out.println("fontの数:" + workbook.getNumberOfFonts());
			// System.out.println("styleの数:" + workbook.getNumCellStyles());
			// System.out.println("dataFormatの数:"
			// + workbook.createDataFormat()
			// .getNumberOfBuiltinBuiltinFormats());

		}

		private void comparePrintSetting() {
			PrintSetup expectedSetting = this.expectedSheet.getPoiSheet()
					.getPrintSetup();
			PrintSetup compareTargetSetting = this.compareTargetSheet
					.getPoiSheet().getPrintSetup();

			assertEquals("Copies: ", expectedSetting.getCopies(),
					compareTargetSetting.getCopies());
			assertEquals("Draft: ", expectedSetting.getDraft(),
					compareTargetSetting.getDraft());
			assertEquals("FooterMargin: ", decimal(expectedSetting
					.getFooterMargin()), decimal(compareTargetSetting
					.getFooterMargin()));
			assertEquals("HeaderMargin: ", decimal(expectedSetting
					.getHeaderMargin()), decimal(compareTargetSetting
					.getHeaderMargin()));
			assertEquals("HResolution: ", expectedSetting.getHResolution(),
					compareTargetSetting.getHResolution());
			assertEquals("Landscape: ", expectedSetting.getLandscape(),
					compareTargetSetting.getLandscape());
			assertEquals("LeftToRight: ", expectedSetting.getLeftToRight(),
					compareTargetSetting.getLeftToRight());
			assertEquals("NoColor: ", expectedSetting.getNoColor(),
					compareTargetSetting.getNoColor());
			assertEquals("NoOrientation: ", expectedSetting.getNoOrientation(),
					compareTargetSetting.getNoOrientation());
			assertEquals("Notes: ", expectedSetting.getNotes(),
					compareTargetSetting.getNotes());
			assertEquals("PageStart: ", expectedSetting.getPageStart(),
					compareTargetSetting.getPageStart());
			assertEquals("PaperSize: ", expectedSetting.getPaperSize(),
					compareTargetSetting.getPaperSize());

			assertEquals("UsePage: ", expectedSetting.getUsePage(),
					compareTargetSetting.getUsePage());
			assertEquals("ValidSettings: ", expectedSetting.getValidSettings(),
					compareTargetSetting.getValidSettings());
			assertEquals("VResolution: ", expectedSetting.getVResolution(),
					compareTargetSetting.getVResolution());
			assertEquals("FitHeight: ", expectedSetting.getFitHeight(),
					compareTargetSetting.getFitHeight());
			assertEquals("FitWidth: ", expectedSetting.getFitWidth(),
					compareTargetSetting.getFitWidth());
			assertEquals("Autobreaks: ", expectedSheet.getPoiSheet()
					.getAutobreaks(), compareTargetSheet.getPoiSheet()
					.getAutobreaks());

			assertEquals("top marging: ", decimal(expectedSheet.getPoiSheet()
					.getMargin(Sheet.TopMargin)),
					decimal(this.compareTargetSheet.getPoiSheet().getMargin(
							Sheet.TopMargin)));

			assertEquals("bottom marging: ", decimal(expectedSheet
					.getPoiSheet().getMargin(Sheet.BottomMargin)),
					decimal(this.compareTargetSheet.getPoiSheet().getMargin(
							Sheet.BottomMargin)));

			assertEquals("left marging: ", decimal(expectedSheet.getPoiSheet()
					.getMargin(Sheet.LeftMargin)),
					decimal(this.compareTargetSheet.getPoiSheet().getMargin(
							Sheet.LeftMargin)));

			assertEquals("right marging: ", decimal(expectedSheet.getPoiSheet()
					.getMargin(Sheet.RightMargin)),
					decimal(this.compareTargetSheet.getPoiSheet().getMargin(
							Sheet.RightMargin)));

			if (!expectedSheet.getPoiSheet().getAutobreaks()) {
				assertEquals("Scale: ", expectedSetting.getScale(),
						compareTargetSetting.getScale());
			}
		}

		private void comareColWidth(int col) {
			assertEquals("列幅の値： " + col + "列目", expectedSheet
					.getColWidthOf(col), this.compareTargetSheet
					.getColWidthOf(col));
		}

		private void compareRowHeight(int row) {
			assertEquals("行高さの値： " + row + "行目", expectedSheet
					.getRowHeightOf(row), this.compareTargetSheet
					.getRowHeightOf(row));
		}

		public void compareValueAndStyleEachOther(ScPos pos) {

			compareCellValueEachOther(pos);
			compareCellStyelEachOther(pos);

			compareCellTypeEachOther(pos);
		}

		private void compareCellValueEachOther(ScPos pos) {

			try {

				ScCell expecetedCell = expectedSheet.getCellAt(pos);

				switch (expecetedCell.getCellType()) {

				case BLANK:
					assertEquals("文字列の値： :" + pos.toReference(), expectedSheet
							.getCellAt(pos).getTextValue(),
							this.compareTargetSheet.getCellAt(pos)
									.getTextValue());
					break;

				case BOOLEAN:
					// TODO 必要になったら実装
					throw new IllegalAccessError("未実装");

				case FORMULA:
					assertEquals("式の値 at " + pos.toReference(), expectedSheet
							.getCellAt(pos).getFormula(),
							this.compareTargetSheet.getCellAt(pos).getFormula());
					break;

				case STRING:

					assertEquals("文字列の値 at :" + pos.toReference(),
							expectedSheet.getCellAt(pos).getTextValue(),
							this.compareTargetSheet.getCellAt(pos)
									.getTextValue());
					break;

				case NUMERIC:

					if (expecetedCell.isDateFormatted()) {
						assertEquals("日付の値 at :" + pos.toReference(),
								expectedSheet

								.getCellAt(pos).getDateValue(),
								this.compareTargetSheet.getCellAt(pos)
										.getDateValue());
					} else {
						assertEquals("数値の値 at :" + pos.toReference(),
								expectedSheet.getCellAt(pos).getNumericValue(),
								this.compareTargetSheet.getCellAt(pos)
										.getNumericValue());
					}
					break;
				case ERROR:
				default:
					break;

				}
			} catch (IllegalAccessError e) {
				throw new IllegalAccessError(e.getMessage() + " at "
						+ pos.toReference());
			} catch (NumberFormatException e) {
				throw new NumberFormatException(e.getMessage() + " at "
						+ pos.toReference());
			}

		}

		private void compareCellTypeEachOther(ScPos pos)
				throws IllegalAccessError {

			ScCell expecetedCell = expectedSheet.getCellAt(pos);
			ScCell compareCell = compareTargetSheet.getCellAt(pos);

			if (expecetedCell.getCellType().isUnknown()) {
				throw new UnsupportedOperationException(
						"未対応のクラスタイプです。 in expected at " + pos.toReference());
			}

			if (compareCell.getCellType().isUnknown()) {
				throw new UnsupportedOperationException(
						"未対応のクラスタイプです。 in compare at " + pos.toReference());
			}

			if (!expecetedCell.getCellType().isBlank()) {
				assertEquals("cellTypeの値： :" + pos.toReference(), expecetedCell
						.getCellType(), compareCell.getCellType());
			}

		}

		private void compareCellStyelEachOther(ScPos pos)
				throws IllegalAccessError {

			assertEquals("スタイル at" + pos.toReference(), expectedSheet
					.getCellAt(pos).getStyle().getStyleWithoutFont(),
					this.compareTargetSheet.getCellAt(pos).getStyle()
							.getStyleWithoutFont());

			assertEquals("フォント at" + pos.toReference(), expectedSheet
					.getCellAt(pos).getStyle().getFontStyle(),
					this.compareTargetSheet.getCellAt(pos).getStyle()
							.getFontStyle());

		}
	}

	public void in(In range) throws FileNotFoundException, IOException {
		this.in(range.start(), range.end());
	}

	public void in(String topLeft, String bottomRight)
			throws FileNotFoundException, IOException {

		if (this.onlyValueAndBorder) {
			compareOnlyValue(topLeft, bottomRight);
		} else {
			compareValueAndStyles(topLeft, bottomRight);
		}

	}

	private void compareOnlyValue(String topLeft, String bottomRight)
			throws FileNotFoundException, IOException {

		ScWorkbook expectedReport;

		expectedReport = TestExpected.fileOf(expectedReportFileName);

		Comparator comp = new Comparator(expectedReport.getFirstSheet(),
				targetSheet);

		comp.compareOnlyValueIn(topLeft, bottomRight);

	}

	private void compareValueAndStyles(String topLeft, String bottomRight)
			throws FileNotFoundException, IOException {
		// TODO 指定範囲の外に値がある場合(validateされない値がある場合)警告する処理を入れる(excel fie
		// systemから読む)

		ScWorkbook expectedReport;
		expectedReport =

		TestExpected.fileOf(expectedReportFileName);

		Comparator comp = new Comparator(expectedReport.getFirstSheet(),
				targetSheet);

		comp.comparaValueAndStyleIn(topLeft, bottomRight);

	}
}
