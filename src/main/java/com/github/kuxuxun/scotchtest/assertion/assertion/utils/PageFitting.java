package com.github.kuxuxun.scotchtest.assertion.assertion.utils;

import org.apache.commons.lang.builder.EqualsBuilder;
import org.apache.commons.lang.builder.HashCodeBuilder;
import org.apache.poi.ss.usermodel.PrintSetup;

import com.github.kuxuxun.scotch.excel.ScSheet;

public class PageFitting {

	private final O o;
	private final int v;
	private final int h;

	private enum O {
		V, H, BOTH, NOT_FIT
	}

	private PageFitting(O o, int v, int h) {
		this.o = o;
		this.v = v;
		this.h = h;
	}

	private static PageFitting factoryOfBothOrVOrH(int v, int h) {
		if (v <= 0) {
			return pageFitToHori(h);
		} else if (h <= 0) {
			return pageFitToVert(v);
		} else {
			return pageFitToBoth(v, h);
		}
	}

	public static PageFitting pageFitToBoth(int v, int h) {
		return new PageFitting(O.BOTH, v, h);
	}

	public static PageFitting pageFitToVert(int v) {
		return new PageFitting(O.V, v, 0);
	}

	public static PageFitting pageFitToHori(int h) {
		return new PageFitting(O.H, 0, h);
	}

	public static PageFitting pageNotFit() {
		return new PageFitting(O.NOT_FIT, 0, 0);
	}

	public static PageFitting getFromSheet(ScSheet sheet) {
		// TODO wb.getPoiSheet().getFitToPage()は関係ある?
		if (!sheet.getPoiSheet().getAutobreaks()) {
			return pageNotFit();
		}

		PrintSetup s = sheet.getPoiSheet().getPrintSetup();
		int v = (s.getFitHeight());
		int h = (s.getFitWidth());

		return factoryOfBothOrVOrH(v, h);

	}

	@Override
	public String toString() {
		StringBuilder b = new StringBuilder();
		b.append("フィットの有無 or 種類: " + o.name() + ", ");
		b.append("縦ページ数 :" + v + ", ");
		b.append("横ページ数 :" + h + " ");

		return b.toString();
	}

	@Override
	public int hashCode() {

		HashCodeBuilder builder = new HashCodeBuilder();
		return builder.append(this.o).append(this.h).append(this.v)
				.toHashCode();
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj) {
			return true;
		}
		if (!(obj instanceof PageFitting)) {
			return false;
		}

		PageFitting that = (PageFitting) obj;

		return new EqualsBuilder().append(this.o, that.o)
				.append(this.h, that.h).append(this.v, that.v).isEquals();
	}
}
