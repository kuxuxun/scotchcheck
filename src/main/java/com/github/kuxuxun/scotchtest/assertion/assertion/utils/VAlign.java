package com.github.kuxuxun.scotchtest.assertion.assertion.utils;

import org.apache.poi.ss.usermodel.CellStyle;

import com.github.kuxuxun.scotch.excel.cell.ScCell;

public enum VAlign {

	VERTICAL_BOTTOM(CellStyle.VERTICAL_BOTTOM), VERTICAL_CENTER(
			CellStyle.VERTICAL_CENTER), VERTICAL_JUSTIFY(
			CellStyle.VERTICAL_JUSTIFY), VERTICAL_TOP(CellStyle.VERTICAL_TOP);

	private short poiAlign;

	private VAlign(short poiAlign) {
		this.poiAlign = poiAlign;
	}

	public static VAlign getFromCell(ScCell cell) {
		CellStyle s = cell.getPoiCell().getCellStyle();
		for (VAlign each : VAlign.values()) {
			if (each.poiAlign == s.getVerticalAlignment()) {
				return each;
			}
		}

		throw new IllegalArgumentException("想定していないAlignです: "
				+ s.getAlignment());
	}

}
