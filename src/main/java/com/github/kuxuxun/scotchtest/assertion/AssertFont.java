package com.github.kuxuxun.scotchtest.assertion;

import static org.junit.Assert.assertEquals;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.cell.ScCell;
import com.github.kuxuxun.scotch.excel.cell.ScColor;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertIn;
import com.github.kuxuxun.scotchtest.assertion.assertion.In;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.FontExample;

public class AssertFont {
	private final ScSheet sheet;

	private AssertFont(ScSheet report) {
		this.sheet = report;
	}

	public static AssertFont that(ScSheet sheet) {
		return new AssertFont(sheet);
	}

	public void fontNameIs(final String expected, In range)
			throws FileNotFoundException, IOException {

		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "のフォント名: ",
						expected, cell.getStyle().getFontStyle().getFontName());
			}
		}.doAssert();

	}

	public void fontSizeIs(final int expected, In range)
			throws FileNotFoundException, IOException {
		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "のフォントサイズ: ",
						expected, cell.getStyle().getFontStyle().getFontSize());
			}
		}.doAssert();

	}

	public void fontColorIs(final ScColor expected, In range)
			throws FileNotFoundException, IOException {
		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "の色: ", expected,
						cell.getStyle().getFontStyle().getFontColor());
			}
		}.doAssert();
	}

	public void fontIsBold(In range) throws FileNotFoundException, IOException {
		assetrBoldIs(true, range);
	}

	public void fontIsNotBold(In range) throws FileNotFoundException,
			IOException {
		assetrBoldIs(false, range);
	}

	private void assetrBoldIs(final boolean positivation, In range)
			throws FileNotFoundException, IOException {
		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "がBold ",
						positivation, cell.getStyle().getFontStyle().isBold());
			}
		}.doAssert();
	}

	public void fontIsItalic(In range) throws FileNotFoundException,
			IOException {
		assertItalicIs(true, range);
	}

	private void assertItalicIs(final boolean positivation, In range)
			throws FileNotFoundException, IOException {
		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "がItalic: ",
						positivation, cell.getStyle().getFontStyle().isItalic());
			}
		}.doAssert();

	}

	public void fontNotItalic(In range) throws FileNotFoundException,
			IOException {
		assertItalicIs(false, range);

	}

	public void fontIsUnderlinedWithSingle(In range)
			throws FileNotFoundException, IOException {
		assertUnderLineSingleIs(true, range);

	}

	public void fontIsNotUnderlinedWithSingle(In range)
			throws FileNotFoundException, IOException {
		assertUnderLineSingleIs(false, range);

	}

	private void assertUnderLineSingleIs(final boolean positivation, In range)
			throws FileNotFoundException, IOException {
		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "が1本下線(会計ではない): ",
						positivation, cell.getStyle().getFontStyle()
								.isUnderlinedWithSingle());
			}
		}.doAssert();
	}

	public void fontIsUnderlinedWithDouble(In range)
			throws FileNotFoundException, IOException {
		this.assertUnderLineDoublleIs(true, range);
	}

	public void fontIsNotUnderlinedWithDouble(In range)
			throws FileNotFoundException, IOException {
		this.assertUnderLineDoublleIs(false, range);
	}

	private void assertUnderLineDoublleIs(final boolean positivation, In range)
			throws FileNotFoundException, IOException {
		new AssertIn(sheet, range) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				assertEquals(cell.getPos().toReference() + "が2本下線(会計ではない): ",
						positivation, cell.getStyle().getFontStyle()
								.isUnderlinedWithDouble());
			}
		}.doAssert();
	}

	public void isSameAs(final FontExample example, In range)
			throws FileNotFoundException, IOException {

		example.doAssert(new AssertFont(sheet), range);

	}

	public ScSheet getTargetSheet() {
		return this.sheet;
	}

}
