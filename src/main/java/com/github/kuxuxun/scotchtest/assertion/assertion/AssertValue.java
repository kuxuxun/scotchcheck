package com.github.kuxuxun.scotchtest.assertion.assertion;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.cell.ScCell;
import com.github.kuxuxun.scotch.excel.cell.ScCell.CellType;

public class AssertValue {

	private final ScSheet sheet;

	private AssertValue(ScSheet sheet) {
		this.sheet = sheet;

	}

	public static AssertValue that(ScSheet sheet) {
		return new AssertValue(sheet);
	}

	public String makeDesc(String description) {
		String desc = "";
		if (n2b(description).length() != 0) {
			desc = description + "として";
		}

		return desc;
	}

	public void isEmpty(final String descritpion, In range) throws Exception {

		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				cellHasExpectedTypeOrInterrupt(cell, CellType.BLANK);

				// 空ならば通る
			}
		}.doAssert();
	}

	public void hasText(final String descritpion, final String expected,
			In range) throws Exception {

		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				cellHasExpectedTypeOrInterrupt(cell, CellType.STRING);

				String desc = makeDesc(descritpion);

				assertEquals(desc + cell.getPos().toReference() + "の値: ",
						expected, cell.getTextValue());
			}
		}.doAssert();
	}

	public void hasDate(final Date expected, In range, final String descritpion)
			throws Exception {

		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				// dateFormattedも見るので cellHasExpectedTypeOrInterruptは使わない
				if (!(cell.getCellType().is(CellType.NUMERIC) && cell
						.isDateFormatted())) {
					throw new IllegalArgumentException(cell.getPos()
							.toReference()
							+ " : の値は日付ではありません");// TODO
					// 独自Exceptionに
				}

				String desc = makeDesc(descritpion);

				assertEquals(desc + cell.getPos().toReference() + "の値: ",
						expected, cell.getDateValue());
			}
		}.doAssert();
	}

	private static String n2b(String s) {
		if (s == null)
			return "";

		return s;
	}

	public void hasNumber(final String descritpion, final BigDecimal expected,
			final int figAferDecimal, In range) throws Exception {

		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				cellHasExpectedTypeOrInterrupt(cell, CellType.NUMERIC);

				String desc = makeDesc(descritpion);

				BigDecimal roundedVal = cell.getNumericValue().setScale(
						figAferDecimal, BigDecimal.ROUND_HALF_UP);

				assertTrue(desc + cell.getPos().toReference() + "の値: expected["
						+ expected.toString() + "] but ["
						+ roundedVal.toString() + "]", expected
						.compareTo(roundedVal) == 0);
			}
		}.doAssert();
	}

	public void hasFormula(final String descritpion, final String expected,
			In range) throws Exception {

		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				cellHasExpectedTypeOrInterrupt(cell, CellType.FORMULA);

				String desc = makeDesc(descritpion);

				assertEquals(desc + cell.getPos().toReference() + "の値: ",
						expected, cell.getFormula());
			}
		}.doAssert();
	}

	public static void cellHasExpectedTypeOrInterrupt(ScCell cell,
			CellType expectedType) {
		if (!cell.getCellType().is(expectedType)) {
			// TODO 独自Exceptionに
			throw new IllegalArgumentException(cell.getPos().toReference()
					+ " : の値は" + expectedType.getDescription() + "ではありません。["
					+ cell.getCellType().getDescription() + "]");

		}
	}

}
