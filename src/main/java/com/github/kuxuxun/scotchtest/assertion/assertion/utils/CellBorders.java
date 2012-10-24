package com.github.kuxuxun.scotchtest.assertion.assertion.utils;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;

import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.cell.ScColor;
import com.github.kuxuxun.scotch.excel.cell.ScStyle;
import com.github.kuxuxun.scotch.excel.cell.ScWithoutFontStyle;

public class CellBorders {

	private final Map<ScPos, CellBorders> matrix;

	private Border onLeft = Border.empty();
	private Border onRight = Border.empty();
	private Border onTop = Border.empty();
	private Border onBottom = Border.empty();
	private final ScPos pos;

	public static CellBorders createOrRegisterdCell(ScPos pos,
			Map<ScPos, CellBorders> matrix) {

		if (matrix.containsKey(pos)) {
			return matrix.get(pos);
		} else {
			return new CellBorders(pos, matrix);
		}

	}

	private CellBorders(ScPos pos, Map<ScPos, CellBorders> matrix) {
		this.pos = pos;
		this.matrix = matrix;
		this.matrix.put(pos, this);
	}

	public Border getOnLeft() {
		return onLeft;
	}

	public void setOnLeft(Border border) {
		if (!border.isAvailable()) {
			return;
		}

		borderOnLeft(border);

		if (!pos.isLeftEdge()) {
			CellBorders neighberLeft = createOrRegisterdCell(this.pos
					.moveInCol(-1), matrix);
			neighberLeft.borderOnRight(this.onLeft);
		}
	}

	public Border getOnRight() {
		return onRight;
	}

	public void setOnRight(Border border) {
		if (!border.isAvailable()) {
			return;
		}
		this.borderOnRight(border);

		// FIXME 右端の場合の処理

		CellBorders neighberRight = createOrRegisterdCell(
				this.pos.moveInCol(1), this.matrix);
		neighberRight.borderOnLeft(this.onRight);
	}

	public Border getOnTop() {
		return onTop;
	}

	public void setOnTop(Border border) {

		if (!border.isAvailable()) {
			return;
		}
		borderOnTop(border);

		if (!pos.isTopEdge()) {
			CellBorders neighberTop = createOrRegisterdCell(this.pos
					.moveInRow(-1), this.matrix);
			neighberTop.borderOnBottom(this.onTop);
		}
	}

	public Border getOnBottom() {
		return onBottom;

	}

	public void setOnBottom(Border border) {
		if (!border.isAvailable()) {
			return;
		}

		// FIXME 下はしの場合の処理
		borderOnBottom(border);
		CellBorders neighberBottom = createOrRegisterdCell(this.pos
				.moveInRow(1), this.matrix);
		neighberBottom.borderOnTop(this.onBottom);
	}

	private void borderOnLeft(Border b) {
		this.onLeft = b;
	}

	private void borderOnRight(Border b) {
		this.onRight = b;
	}

	private void borderOnTop(Border b) {
		this.onTop = b;
	}

	private void borderOnBottom(Border b) {
		this.onBottom = b;
	}

	public ScPos getCellPosition() {
		return this.pos;
	}

	public static class Border {

		public static Border none(ScColor color) {
			return Border.as(CellStyle.BORDER_NONE, color);
		}

		public static Border thin(ScColor color) {
			return Border.as(CellStyle.BORDER_THIN, color);
		}

		public static Border medium(ScColor color) {
			return Border.as(CellStyle.BORDER_MEDIUM, color);
		}

		public static Border dashed(ScColor color) {
			return Border.as(CellStyle.BORDER_DASHED, color);
		}

		public static Border hair(ScColor color) {
			return Border.as(CellStyle.BORDER_HAIR, color);
		}

		public static Border thick(ScColor color) {
			return Border.as(CellStyle.BORDER_THICK, color);
		}

		public static Border doubleLine(ScColor color) {
			return Border.as(CellStyle.BORDER_DOUBLE, color);
		}

		public static Border dotted(ScColor color) {
			return Border.as(CellStyle.BORDER_DOTTED, color);
		}

		public static Border mediumDashed(ScColor color) {
			return Border.as(CellStyle.BORDER_MEDIUM_DASHED, color);
		}

		public static Border dashDot(ScColor color) {
			return Border.as(CellStyle.BORDER_DASH_DOT, color);
		}

		public static Border mediumDash_Dot(ScColor color) {
			return Border.as(CellStyle.BORDER_MEDIUM_DASH_DOT, color);
		}

		public static Border dashDotDot(ScColor color) {
			return Border.as(CellStyle.BORDER_DASH_DOT_DOT, color);
		}

		public static Border mediumDashDotDot(ScColor color) {
			return Border.as(CellStyle.BORDER_MEDIUM_DASH_DOT_DOT, color);
		}

		public static Border slantedDashDot(ScColor color) {
			return Border.as(CellStyle.BORDER_SLANTED_DASH_DOT, color);
		}

		private short borderStyle = -1;
		private short color = -1;

		public static Border empty() {
			return new Border((short) -1, (short) -1);
		}

		public static Border thinBlack() {
			return Border.as(CellStyle.BORDER_THIN, ScColor.BLACK);
		}

		public static Border thinAutoColor() {
			return Border.as(CellStyle.BORDER_THIN, ScColor.NON_OR_AUTO);
		}

		public static Border as(short style, ScColor color) {
			return new Border(style, color.getIndex());
		}

		public static Border as(short style, short color) {
			return new Border(style, color);
		}

		public Border(short style, short color) {
			this.borderStyle = style;
			this.color = color;
		}

		public short getBorderStyle() {
			return borderStyle;
		}

		public void setBorderStyle(short borderStyle) {
			this.borderStyle = borderStyle;
		}

		public short getColor() {
			return color;
		}

		public void setColor(short color) {
			this.color = color;
		}

		public boolean isAvailable() {
			return (color > 0 && borderStyle > 0);
		}

		@Override
		public boolean equals(Object o) {
			if (this == o) {
				return true;
			}

			if (!(o instanceof Border)) {
				return false;
			}

			Border counter = (Border) o;

			if (this.borderStyle == counter.borderStyle
					&& this.color == counter.color) {
				return true;
			}

			return false;
		}

		@Override
		public int hashCode() {
			int result = borderStyle;
			result = result * 31 + color;
			return result;
		}

		@Override
		public String toString() {
			return "[style:" + borderStyle + ",color:" + color + "]";
		}

		public String formatOfTuppleWith(Border c) {
			return "[(" + this.borderStyle + "," + this.color + ")," + "("
					+ c.borderStyle + "," + c.color + ")]";
		}
	}

	public CellBorders getLeftNeighberCell() {
		return this.matrix.get(this.pos.moveInCol(-1));
	}

	public CellBorders getRightNeighberCell() {
		return this.matrix.get(this.pos.moveInCol(1));
	}

	public CellBorders getTopNeighberCell() {
		return this.matrix.get(this.pos.moveInRow(-1));
	}

	public CellBorders getBottomNeighberCell() {
		return this.matrix.get(this.pos.moveInRow(1));
	}

	public void getFromCellStyle(ScStyle style) {

		ScWithoutFontStyle borderStyle = style.getStyleWithoutFont();

		Border left = Border.as(borderStyle.getBorderLeft(), borderStyle
				.getLeftBorderColor());

		this.setOnLeft(left);

		Border right = Border.as(borderStyle.getBorderRight(), borderStyle
				.getRightBorderColor());

		this.setOnRight(right);

		Border top = Border.as(borderStyle.getBorderTop(), borderStyle
				.getTopBorderColor());

		this.setOnTop(top);

		Border bottom = Border.as(borderStyle.getBorderBottom(), borderStyle
				.getBottomBorderColor());

		this.setOnBottom(bottom);
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) {
			return true;
		}

		if (!(o instanceof CellBorders)) {
			return false;
		}

		CellBorders counter = (CellBorders) o;

		if (!this.pos.equals(counter.pos)) {
			return false;
		}

		if (!this.onLeft.equals(counter.getOnLeft())) {
			return false;
		}

		if (!this.onRight.equals(counter.getOnRight())) {
			return false;
		}

		if (!this.onTop.equals(counter.getOnTop())) {
			return false;
		}

		if (!this.onLeft.equals(counter.getOnLeft())) {
			return false;
		}

		return true;
	}

	@Override
	public int hashCode() {
		int result = pos.hashCode();
		result = result * 31 + onLeft.hashCode();
		result = result * 31 + onRight.hashCode();
		result = result * 31 + onTop.hashCode();
		result = result * 31 + onBottom.hashCode();
		return result;
	}

	@Override
	public String toString() {

		String result = "[" + this.pos.toReference() + ":,";
		result += "left : " + onLeft.toString();
		result += "right : " + onRight.toString();
		result += "top : " + onTop.toString();
		result += "bottom : " + onBottom.toString();
		result += "]";

		return result;
	}

	public String getDiffrenceWith(CellBorders cell) {
		if (this.equals(cell)) {
			return "";
		}

		if (!pos.equals(cell.pos)) {
			return "difference cell [" + this.pos.toReference() + ","
					+ cell.pos.toReference() + "]";
		}

		String result = "";
		if (!onLeft.equals(cell.onLeft)) {
			result += " left:" + onLeft.formatOfTuppleWith(cell.onLeft);
		}

		if (!onRight.equals(cell.onRight)) {
			result += " right:" + onRight.formatOfTuppleWith(cell.onRight);
		}
		if (!onBottom.equals(cell.onBottom)) {
			result += " bottom:" + onBottom.formatOfTuppleWith(cell.onBottom);
		}
		if (!onTop.equals(cell.onTop)) {
			result += " top:" + onTop.formatOfTuppleWith(cell.onTop);
		}

		return (result.length() == 0 ? "" : "at " + pos.toReference() + result);
	}

	public static CellBorders bordersToCellBorder(final Border left,
			final Border top, final Border right, final Border bottom) {

		CellBorders e = CellBorders.createOrRegisterdCell(new ScPos("A1"),
				new HashMap<ScPos, CellBorders>());

		e.setOnLeft(left);
		e.setOnTop(top);
		e.setOnRight(right);
		e.setOnBottom(bottom);
		return e;
	}

}
