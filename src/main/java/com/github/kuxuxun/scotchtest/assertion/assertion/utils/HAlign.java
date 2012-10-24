package com.github.kuxuxun.scotchtest.assertion.assertion.utils;

import org.apache.poi.ss.usermodel.CellStyle;

import com.github.kuxuxun.scotch.excel.cell.ScCell;

public enum HAlign {

	GENERAL(CellStyle.ALIGN_GENERAL), LEFT(CellStyle.ALIGN_LEFT), CENTER(
			CellStyle.ALIGN_CENTER), RIGHT(CellStyle.ALIGN_RIGHT), FILL(
			CellStyle.ALIGN_FILL), JUSTIFY(CellStyle.ALIGN_JUSTIFY), CENTER_SELECTION(
			CellStyle.ALIGN_CENTER_SELECTION);

	private short poiAlign;

	private HAlign(short poiAlign) {
		this.poiAlign = poiAlign;
	}

	public static HAlign getFromCell(ScCell cell) {
		CellStyle s = cell.getPoiCell().getCellStyle();
		for (HAlign each : HAlign.values()) {
			if (each.poiAlign == s.getAlignment()) {
				return each;
			}
		}

		throw new IllegalArgumentException("想定していないAlignです: "
				+ s.getAlignment());
	}

}
