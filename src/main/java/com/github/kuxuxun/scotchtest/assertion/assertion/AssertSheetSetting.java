package com.github.kuxuxun.scotchtest.assertion.assertion;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import com.github.kuxuxun.scotch.excel.ScSheet;

public class AssertSheetSetting {

	private static String n2b(String s) {
		if (s == null)
			return "";

		return s;
	}

	private final ScSheet sheet;

	private AssertSheetSetting(ScSheet sheet) {
		this.sheet = sheet;
	}

	public static AssertSheetSetting that(ScSheet sheet) {
		return new AssertSheetSetting(sheet);
	}

	public String makeDesc(String description) {
		String desc = "";
		if (n2b(description).length() != 0) {
			desc = description + "として";
		}

		return desc;
	}

	public void isProtected() {

		assertTrue("シートが保護されている", sheet.isProtected());

	}

	public void isNotProtected() {

		assertFalse("シートが保護されていない", sheet.isProtected());

	}

}
