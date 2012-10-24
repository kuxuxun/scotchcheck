package com.github.kuxuxun.scotchtest.assertion.assertion.utils;

public enum DataFormatExample {

	/** e.g 1000 => 1000 */
	INTEGER_NEGRED("0_ ;[Red]\\-0\\ "),

	/** e.g 1000.5 => 1,001 */
	INTEGER_COMMA_NEGRED("#,###;[Red]\\-#,###"),

	/** e.g 1000.5 => 1,001 */
	INTEGER_COMMA_NEGRED_LAST_0("#,##0_ ;[Red]\\-#,##0\\ "),

	/** e.g 2000500 => 2,001 */
	INTEGER_IN_KILO_NEGRED("#,##0,;[Red]\\-#,##0,;#"),

	/** e.g 200.25 => 200.3% */
	PERCENT_ROUND_AT_SEC_NEGRED("#,##0.0\\%;[Red]\\-#,##0.0\\%;#"),

	/** e.g 200.25 => 200.25 */
	INTEGER_AFTER_POINT_TWO_NEGRED("#,##0.00_ ;[Red]\\-#,##0.00\\ "),

	/** 文字列 */
	TEXT("@"),
	/** 標準 */
	GENERAL("General");

	private String formatString;

	private DataFormatExample(String format) {
		this.formatString = format;
	}

	public String getFormatString() {
		return formatString;
	}
}
