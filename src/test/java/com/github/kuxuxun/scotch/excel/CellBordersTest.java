package com.github.kuxuxun.scotch.excel;

import static org.hamcrest.core.Is.is;
import static org.junit.Assert.assertThat;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.junit.Before;
import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestHelper;
import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.cell.ScColor;
import com.github.kuxuxun.scotch.excel.cell.ScStyle;
import com.github.kuxuxun.scotch.excel.cell.ScWithoutFontStyle;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.CellBorders;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.CellBorders.Border;

public class CellBordersTest extends TestHelper {

	private CellBorders b;
	private Map<ScPos, CellBorders> cellMatrix;

	@Before
	public void setup() {

		cellMatrix = new HashMap<ScPos, CellBorders>();
		b = CellBorders.createOrRegisterdCell(new ScPos("C3"), cellMatrix);
	}

	@Test
	public void testGetDifference() {
		b.setOnLeft(Border.thinBlack());

		CellBorders other = CellBorders.createOrRegisterdCell(new ScPos("C3"),
				new HashMap<ScPos, CellBorders>());

		other.setOnTop(Border.thinBlack());

		assertThat(b.getDiffrenceWith(other),
				is("at C3 left:[(1,8),(-1,-1)] top:[(-1,-1),(1,8)]"));

	}

	@Test
	public void testUpdateNeighberOnCraete() {
		b.setOnLeft(Border.thinAutoColor());

		CellBorders neighber = CellBorders.createOrRegisterdCell(
				new ScPos("B3"), cellMatrix);

		neighber.setOnTop(Border.thinAutoColor());

		assertThat(neighber.getOnRight(), is(Border.thinAutoColor()));
		assertThat(neighber.getOnTop(), is(Border.thinAutoColor()));

	}

	@Test
	public void testUpdateNeighberOnSetBorder() {
		b.setOnRight(Border.thinAutoColor());

		CellBorders stepNeighber = CellBorders.createOrRegisterdCell(new ScPos(
				"E3"), cellMatrix);

		stepNeighber.setOnLeft(Border.thinAutoColor());

		CellBorders neighber = cellMatrix.get(new ScPos("D3"));

		assertThat(neighber.getOnRight(), is(Border.thinAutoColor()));
		assertThat(neighber.getOnLeft(), is(Border.thinAutoColor()));

	}

	@Test
	public void testCreateNeighberCellOnSetBorderOnLeft() throws Exception {

		b.setOnLeft(Border.as(HSSFCellStyle.BORDER_THIN, ScColor.BLACK));

		CellBorders cell = b.getLeftNeighberCell();

		assertThat(cell.getOnRight(), is(Border.as(HSSFCellStyle.BORDER_THIN,
				ScColor.BLACK)));

		assertThat(cell.getCellPosition().toReference(), is("B3"));

	}

	@Test
	public void testCreateNeighberCellOnSetBorderOnRight() throws Exception {

		b.setOnRight(Border.as(HSSFCellStyle.BORDER_THIN, ScColor.BLACK));

		CellBorders cell = b.getRightNeighberCell();

		assertThat(cell.getOnLeft(), is(Border.as(HSSFCellStyle.BORDER_THIN,
				ScColor.BLACK)));

		assertThat(cell.getCellPosition().toReference(), is("D3"));

	}

	@Test
	public void testCreateNeighberCellOnSetBorderOnTop() throws Exception {

		b.setOnTop(Border.as(HSSFCellStyle.BORDER_THIN, ScColor.BLACK));

		CellBorders cell = b.getTopNeighberCell();

		assertThat(cell.getOnBottom(), is(Border.as(HSSFCellStyle.BORDER_THIN,
				ScColor.BLACK)));

		assertThat(cell.getCellPosition().toReference(), is("C2"));

	}

	@Test
	public void testCreateNeighberCellOnSetBorderOnBottom() throws Exception {

		b.setOnBottom(Border.as(HSSFCellStyle.BORDER_THIN, ScColor.BLACK));

		CellBorders cell = b.getBottomNeighberCell();

		assertThat(cell.getOnTop(), is(Border.thinBlack()));

		assertThat(cell.getCellPosition().toReference(), is("C4"));

	}

	@Test
	public void testSetBorderOnLeft() throws Exception {

		b.setOnLeft(Border.thinAutoColor());

		assertThat(b.getOnLeft(), is(Border.thinAutoColor()));
	}

	@Test
	public void testSetBorderOnRight() throws Exception {

		b.setOnRight(Border.thinBlack());

		assertThat(b.getOnRight(), is(Border.as(HSSFCellStyle.BORDER_THIN,
				ScColor.BLACK)));
	}

	@Test
	public void testSetBorderOnTop() throws Exception {

		b.setOnTop(Border.thinBlack());

		assertThat(b.getOnTop(), is(Border.as(HSSFCellStyle.BORDER_THIN,
				ScColor.BLACK)));
	}

	@Test
	public void testSetBorderOnBottom() throws Exception {

		b.setOnBottom(Border.thinAutoColor());

		assertThat(b.getOnBottom(), is(Border.thinAutoColor()));
	}

	@Test
	public void testBorderIsNotRemovable() throws IllegalAccessException {

		b.setOnLeft(Border.thinAutoColor());
		b.setOnRight(Border.thinAutoColor());
		b.setOnTop(Border.thinAutoColor());
		b.setOnBottom(Border.thinAutoColor());

		b.setOnLeft(Border.empty());
		b.setOnRight(Border.empty());
		b.setOnTop(Border.empty());
		b.setOnBottom(Border.empty());

		assertThat(b.getOnLeft(), is(Border.thinAutoColor()));
		assertThat(b.getOnRight(), is(Border.thinAutoColor()));
		assertThat(b.getOnBottom(), is(Border.thinAutoColor()));
		assertThat(b.getOnTop(), is(Border.thinAutoColor()));
	}

	@Test
	public void testSetByCelStyle() throws Exception {

		ScStyle style = createMockCellStyle();

		b.getFromCellStyle(style);

		assertThat(b.getOnLeft(), is(Border.thinBlack()));
		assertThat(b.getOnRight(), is(Border.thinBlack()));
		assertThat(b.getOnBottom(), is(Border.thinBlack()));
		assertThat(b.getOnTop(), is(Border.thinBlack()));
	}

	private ScStyle createMockCellStyle() throws IllegalAccessException {
		ScWithoutFontStyle styleWithoutFont = ScWithoutFontStyle
				.createWithPoiStyle(new MockCellStyle());

		styleWithoutFont.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		styleWithoutFont.setLeftBorderColor(new HSSFColor.BLACK().getIndex());

		styleWithoutFont.setBorderRight(HSSFCellStyle.BORDER_THIN);
		styleWithoutFont.setRightBorderColor(new HSSFColor.BLACK().getIndex());

		styleWithoutFont.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		styleWithoutFont.setBottomBorderColor(new HSSFColor.BLACK().getIndex());

		styleWithoutFont.setBorderTop(HSSFCellStyle.BORDER_THIN);
		styleWithoutFont.setTopBorderColor(new HSSFColor.BLACK().getIndex());

		ScStyle style = ScStyle.createNew(decimal(0), decimal(0));
		setValueToFieldOf(style, "cellStyleWithoutFont", styleWithoutFont);
		return style;
	}

	public class MockCellStyle implements CellStyle {

		public void cloneStyleFrom(CellStyle source) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public short getAlignment() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getBorderBottom() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getBorderLeft() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getBorderRight() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getBorderTop() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getBottomBorderColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getDataFormat() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public String getDataFormatString() {
			// TODO 自動生成されたメソッド・スタブ
			return null;
		}

		public short getFillBackgroundColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public Color getFillBackgroundColorColor() {
			// TODO 自動生成されたメソッド・スタブ
			return null;
		}

		public short getFillForegroundColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public Color getFillForegroundColorColor() {
			// TODO 自動生成されたメソッド・スタブ
			return null;
		}

		public short getFillPattern() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getFontIndex() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public boolean getHidden() {
			// TODO 自動生成されたメソッド・スタブ
			return false;
		}

		public short getIndention() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getIndex() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getLeftBorderColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public boolean getLocked() {
			// TODO 自動生成されたメソッド・スタブ
			return false;
		}

		public short getRightBorderColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getRotation() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getTopBorderColor() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public short getVerticalAlignment() {
			// TODO 自動生成されたメソッド・スタブ
			return 0;
		}

		public boolean getWrapText() {
			// TODO 自動生成されたメソッド・スタブ
			return false;
		}

		public void setAlignment(short align) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setBorderBottom(short border) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setBorderLeft(short border) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setBorderRight(short border) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setBorderTop(short border) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setBottomBorderColor(short color) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setDataFormat(short fmt) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setFillBackgroundColor(short bg) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setFillForegroundColor(short bg) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setFillPattern(short fp) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setFont(Font font) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setHidden(boolean hidden) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setIndention(short indent) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setLeftBorderColor(short color) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setLocked(boolean locked) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setRightBorderColor(short color) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setRotation(short rotation) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setTopBorderColor(short color) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setVerticalAlignment(short align) {
			// TODO 自動生成されたメソッド・スタブ

		}

		public void setWrapText(boolean wrapped) {
			// TODO 自動生成されたメソッド・スタブ

		}

	}
}
