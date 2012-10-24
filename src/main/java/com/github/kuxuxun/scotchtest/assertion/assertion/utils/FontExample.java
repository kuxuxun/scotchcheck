package com.github.kuxuxun.scotchtest.assertion.assertion.utils;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.github.kuxuxun.scotch.excel.cell.ScColor;
import com.github.kuxuxun.scotchtest.assertion.AssertFont;
import com.github.kuxuxun.scotchtest.assertion.assertion.In;

public class FontExample {

	private final List<TestableAttribute<?>> attributes = new ArrayList<TestableAttribute<?>>();

	private FontExample(Builder builder) {
		attributes.add(builder.underLineSingle);
		attributes.add(builder.underLineDouble);
		attributes.add(builder.italic);
		attributes.add(builder.bold);
		attributes.add(builder.fontName);
		attributes.add(builder.size);
		attributes.add(builder.color);
	}

	@Override
	public String toString() {
		StringBuffer sb = new StringBuffer();
		sb.append(" [FontExamle ");
		for (TestableAttribute<?> each : attributes) {
			if (each.isApplicantToTest()) {
				sb.append(" " + each.getWhatAbout().name() + "=\""
						+ each.value.toString() + "\" ");
			}
		}

		return sb.toString();
	}

	private static abstract class TestableAttribute<T> {
		TestApplicants whatAbout;
		private T value;
		private boolean testApplication;

		public TestableAttribute(TestApplicants ta) {
			whatAbout = ta;
		}

		public T getExpectedValue() {
			return value;
		}

		public void setValue(T value) {
			this.value = value;
			setTestApplication(true);
		}

		public boolean isApplicantToTest() {
			return testApplication;
		}

		public void setTestApplication(boolean testApplication) {
			this.testApplication = testApplication;
		}

		public TestApplicants getWhatAbout() {
			return whatAbout;
		}

		protected abstract void assertion(In range, AssertFont assertFont)
				throws FileNotFoundException, IOException;

		public void doAssertIfApplicant(In range, AssertFont assertFont)
				throws FileNotFoundException, IOException {
			if (isApplicantToTest()) {
				assertion(range, assertFont);
			}
		}

	}

	private enum TestApplicants {
		UNDER_LINE_SINGLE, UNDER_LINE_DOUBLE, ITALIC, BOLD, FONT_NAME, FONT_SIZE, COLOR;
	}

	public static class Builder {

		private final TestableAttribute<ScColor> color = new TestableAttribute<ScColor>(
				TestApplicants.COLOR) {
			@Override
			public void assertion(In range, AssertFont assertFont)
					throws FileNotFoundException, IOException {

				assertFont.fontColorIs(getExpectedValue(), range);
			}
		};

		private final TestableAttribute<Boolean> underLineSingle = new TestableAttribute<Boolean>(
				TestApplicants.UNDER_LINE_SINGLE) {
			@Override
			public void assertion(In range, AssertFont assertFont)
					throws FileNotFoundException, IOException {

				if (getExpectedValue() == true) {
					assertFont.fontIsUnderlinedWithSingle(range);
				} else {
					assertFont.fontIsNotUnderlinedWithSingle(range);
				}

			}
		};

		private final TestableAttribute<Boolean> underLineDouble = new TestableAttribute<Boolean>(
				TestApplicants.UNDER_LINE_DOUBLE) {

			@Override
			public void assertion(In range, AssertFont assertFont)
					throws FileNotFoundException, IOException {

				if (getExpectedValue() == true) {
					assertFont.fontIsUnderlinedWithDouble(range);
				} else {
					assertFont.fontIsNotUnderlinedWithDouble(range);
				}

			}

		};
		private final TestableAttribute<Boolean> italic = new TestableAttribute<Boolean>(
				TestApplicants.ITALIC) {

			@Override
			public void assertion(In range, AssertFont assertFont)
					throws FileNotFoundException, IOException {
				if (getExpectedValue() == true) {
					assertFont.fontIsItalic(range);
				} else {
					assertFont.fontNotItalic(range);
				}
			}
		};
		private final TestableAttribute<Boolean> bold = new TestableAttribute<Boolean>(
				TestApplicants.BOLD) {

			@Override
			public void assertion(In range, AssertFont assertFont)
					throws FileNotFoundException, IOException {
				if (getExpectedValue() == true) {
					assertFont.fontIsBold(range);
				} else {
					assertFont.fontIsNotBold(range);
				}
			}
		};
		private final TestableAttribute<String> fontName = new TestableAttribute<String>(
				TestApplicants.FONT_NAME) {

			@Override
			public void assertion(In range, AssertFont assertFont)
					throws FileNotFoundException, IOException {
				assertFont.fontNameIs(getExpectedValue(), range);

			}
		};
		private final TestableAttribute<Integer> size = new TestableAttribute<Integer>(
				TestApplicants.FONT_SIZE) {

			@Override
			public void assertion(In range, AssertFont assertFont)
					throws FileNotFoundException, IOException {
				assertFont.fontSizeIs(getExpectedValue(), range);
			}
		};

		public Builder setColor(ScColor color) {
			this.color.setValue(color);
			return this;
		}

		public static Builder simpleArial(int size) {
			return new Builder().setFontName("Arial").setSize(size);
		}

		public static Builder simpleMSPGothic(int size) {
			return new Builder().setFontName("ＭＳ Ｐゴシック").setSize(size);
		}

		public static Builder simpleMSGothic(int size) {
			return new Builder().setFontName("ＭＳ ゴシック").setSize(size);
		}

		public FontExample build() {
			return new FontExample(this);
		}

		public Builder setUnderLineSingle(boolean underLineSingle) {
			this.underLineSingle.setValue(underLineSingle);
			return this;
		}

		public Builder setItalic(boolean italic) {
			this.italic.setValue(italic);
			return this;
		}

		public Builder setBold(boolean bold) {
			this.bold.setValue(bold);
			return this;
		}

		public Builder setFontName(String fontName) {
			this.fontName.setValue(fontName);
			return this;
		}

		public Builder setSize(int size) {
			this.size.setValue(size);
			return this;
		}

		public Builder setUnderLineDouble(boolean underLinaeDouble) {
			this.underLineDouble.setValue(underLinaeDouble);
			return this;
		}
	}

	public void doAssert(final AssertFont assertFont, In range)
			throws FileNotFoundException, IOException {
		for (TestableAttribute<?> each : attributes) {
			each.doAssertIfApplicant(range, assertFont);
		}
	}

}
