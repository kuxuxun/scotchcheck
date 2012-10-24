package com.github.kuxuxun.scotchtest.assertion.assertion.utils;

import org.apache.commons.lang.builder.EqualsBuilder;
import org.apache.commons.lang.builder.HashCodeBuilder;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;

import com.github.kuxuxun.scotch.excel.ScSheet;

public class PageNumberPrint {

	private static class Setting {

		private final String format;

		private Setting(String format) {
			this.format = format;
		}

		public String getFormat() {
			return format;
		}

		@Override
		public boolean equals(Object o) {
			if (!(o instanceof Setting)) {
				return false;
			}

			Setting that = (Setting) o;
			return this.format.equals(that.getFormat());

		}

		@Override
		public int hashCode() {
			return this.format.hashCode();
		}

		@Override
		public String toString() {
			return this.format;
		}
	}

	private Setting hLeft, hCenter, hRight;
	private Setting fLeft, fCenter, fRight;

	private PageNumberPrint() {
		hLeft = new Setting("");
		hCenter = new Setting("");
		hRight = new Setting("");
		fLeft = new Setting("");
		fCenter = new Setting("");
		fRight = new Setting("");
	}

	public static PageNumberPrint pageOfAllAtCenterFooter(String delimiter,
			String prefix) {
		String format = MarkupTag.PAGE_FIELD.getRepresentation() + delimiter
				+ MarkupTag.NUM_PAGES_FIELD.getRepresentation() + prefix;

		PageNumberPrint p = new PageNumberPrint();

		p.fCenter = new Setting(format);

		return p;

	}

	private enum MarkupTag {
		SHEET_NAME_FIELD("&A", false), DATE_FIELD("&D", false), FILE_FIELD(
				"&F", false), FULL_FILE_FIELD("&Z", false), PAGE_FIELD("&P",
				false), TIME_FIELD("&T", false), NUM_PAGES_FIELD("&N", false),

		PICTURE_FIELD("&G", false),

		BOLD_FIELD("&B", true), ITALIC_FIELD("&I", true), STRIKETHROUGH_FIELD(
				"&S", true), SUBSCRIPT_FIELD("&Y", true), SUPERSCRIPT_FIELD(
				"&X", true), UNDERLINE_FIELD("&U", true), DOUBLE_UNDERLINE_FIELD(
				"&E", true), ;

		private final String _representation;
		private final boolean _occursInPairs;

		private MarkupTag(String sequence, boolean occursInPairs) {
			_representation = sequence;
			_occursInPairs = occursInPairs;
		}

		public String getRepresentation() {
			return _representation;
		}

		public boolean occursPairs() {
			return _occursInPairs;
		}
	}

	public static PageNumberPrint getFromSheet(ScSheet wb) {

		PageNumberPrint p = new PageNumberPrint();

		Header h = wb.getPoiSheet().getHeader();
		p.setHLeft(new Setting(h.getLeft()));
		p.setHCenter(new Setting(h.getCenter()));
		p.setHRight(new Setting(h.getRight()));

		Footer f = wb.getPoiSheet().getFooter();

		p.setFLeft(new Setting(f.getLeft()));
		p.setFCenter(new Setting(f.getCenter()));
		p.setFRight(new Setting(f.getRight()));

		return p;
	}

	@Override
	public int hashCode() {

		HashCodeBuilder builder = new HashCodeBuilder();
		builder.append(this.hLeft).append(this.hCenter).append(this.hRight);
		builder.append(this.fLeft).append(this.fCenter).append(this.fRight);
		return builder.toHashCode();

	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj) {
			return true;
		}
		if (!(obj instanceof PageNumberPrint)) {
			return false;
		}

		PageNumberPrint that = (PageNumberPrint) obj;

		EqualsBuilder builder = new EqualsBuilder();
		builder.append(this.hLeft, that.hLeft).append(this.hCenter,
				that.hCenter).append(this.hRight, that.hRight);

		builder.append(this.fLeft, that.fLeft).append(this.fCenter,
				that.fCenter).append(this.fRight, that.fRight);

		return builder.isEquals();
	}

	@Override
	public String toString() {
		StringBuilder b = new StringBuilder();
		b.append(" header left :" + hLeft.getFormat());
		b.append(" header center :" + hCenter.getFormat());
		b.append(" header right :" + hRight.getFormat());

		b.append(" footer left :" + fLeft.getFormat());
		b.append(" footer center :" + fCenter.getFormat());
		b.append(" footer right :" + fRight.getFormat());

		return b.toString();
	}

	public Setting getHLeft() {
		return hLeft;
	}

	public void setHLeft(Setting left) {
		hLeft = left;
	}

	public Setting getHCenter() {
		return hCenter;
	}

	public void setHCenter(Setting center) {
		hCenter = center;
	}

	public Setting getHRight() {
		return hRight;
	}

	public void setHRight(Setting right) {
		hRight = right;
	}

	public Setting getFLeft() {
		return fLeft;
	}

	public void setFLeft(Setting left) {
		fLeft = left;
	}

	public Setting getFCenter() {
		return fCenter;
	}

	public void setFCenter(Setting center) {
		fCenter = center;
	}

	public Setting getFRight() {
		return fRight;
	}

	public void setFRight(Setting right) {
		fRight = right;
	}
}
