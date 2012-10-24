package com.github.kuxuxun.scotchtest.assertion;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.cell.ScCell;
import com.github.kuxuxun.scotch.excel.cell.ScColor;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertIn;
import com.github.kuxuxun.scotchtest.assertion.assertion.At;
import com.github.kuxuxun.scotchtest.assertion.assertion.In;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.BordersInSheet;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.CellBorders;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.CellBorders.Border;

//TODO 重複大杉なんでリファクタ
public class AssertBorders {

	private final ScSheet sheet;

	private AssertBorders(ScSheet report) {
		this.sheet = report;

	}

	public static AssertBorders that(ScSheet file) {
		return new AssertBorders(file);

	}

	public void isSurroundByBorderWith(Border border, In range)
			throws FileNotFoundException, IOException {

		isSurroundByBorderWith(border, border, border, border, range);
	}

	public void isSurroundByBoxOfThinBorder(final In range)
			throws FileNotFoundException, IOException {

		isSurroundInBoxWith(Border.thin(ScColor.NON_OR_AUTO), range);

	}

	public void isSurroundByBorderWith(final Border left, final Border top,
			final Border right, final Border bottom, final In range)
			throws FileNotFoundException, IOException {

		CellBorders e = CellBorders.bordersToCellBorder(left, top, right,
				bottom);

		isSurroundByBorderWith(e, range);
	}

	public void isSurroundByBorderWith(final CellBorders expectBorderStyle,
			final In range) throws FileNotFoundException, IOException {
		new Assert(sheet, range) {

			@Override
			public void withBorders(ScCell cell, CellBorders bordersOfTheCell,
					BordersInSheet borders) {

				ScPos pos = cell.getPos();

				if (range.isLeftEdge(pos)) {
					assertLeftBorder(pos, bordersOfTheCell, expectBorderStyle
							.getOnLeft());
				}

				if (range.isTopEdge(pos)) {
					assertTopBorder(pos, bordersOfTheCell, expectBorderStyle
							.getOnTop());
				}

				if (range.isRightEdge(pos)) {
					assertRightBorder(pos, bordersOfTheCell, expectBorderStyle
							.getOnRight());
				}

				if (range.isBottomEdge(pos)) {
					assertBottomBorder(pos, bordersOfTheCell, expectBorderStyle
							.getOnBottom());
				}

			}

		}.doAssert();

	}

	public void eachCellsSurroundByThinBorder(final In range) throws Exception {
		eachCellsSurroundBy(Border.thinAutoColor(), range);
	}

	public void eachCellsSurroundBy(final Border expectBorder, final In range)
			throws Exception {

		eachCellsSurroundBy(expectBorder, expectBorder, expectBorder,
				expectBorder, range);

	}

	public void eachCellsSurroundBy(final Border left, final Border top,
			final Border right, final Border bottom, final In range)
			throws Exception {

		CellBorders e = CellBorders.bordersToCellBorder(left, top, right,
				bottom);

		eachCellsSurroundBy(e, range);

	}

	public void eachCellsSurroundBy(final CellBorders expectBorders,
			final In range) throws Exception {

		new Assert(sheet, range) {

			@Override
			public void withBorders(ScCell cell, CellBorders bordersOfTheCell,
					BordersInSheet borders) throws FileNotFoundException,
					IOException {
				isSurroundByBorderWith(expectBorders, At.at(cell.getPos()));

			}

		}.doAssert();
	}

	private void assertBottomBorder(ScPos pos, CellBorders cb,
			Border expectBorder) {
		assertEquals(pos.toReference() + "の下の罫線", expectBorder, cb
				.getOnBottom());
	}

	private void assertRightBorder(ScPos pos, CellBorders cb,
			Border expectBorder) {
		assertEquals(pos.toReference() + "の右の罫線", expectBorder, cb.getOnRight());
	}

	private void assertTopBorder(ScPos pos, CellBorders cb, Border expect) {
		assertEquals(pos.toReference() + "の上の罫線", expect, cb.getOnTop());
	}

	private void assertLeftBorder(ScPos pos, CellBorders cb, Border expect) {
		assertEquals(pos.toReference() + "の左の罫線", expect, cb.getOnLeft());
	}

	private void assertAroundBoxIfApropriateCell(ScPos pos,
			CellBorders bordersInRange, CellBorders expectBorders, In range) {

		assertLeftBorderOnlyIfEdge(pos, bordersInRange, expectBorders
				.getOnLeft(), range);

		assertRightBorderOnlyIfEdge(pos, bordersInRange, expectBorders
				.getOnRight(), range);

		assertTopBorderOnlyIfEdge(pos, bordersInRange,
				expectBorders.getOnTop(), range);

		assertBottomBorderOnlyIfEdge(pos, bordersInRange, expectBorders
				.getOnBottom(), range);
	}

	private void assertNoInternalLineIfApropriateCell(ScPos pos,
			CellBorders bordersInRange, In range) {

		if (!range.isLeftEdge(pos)) {
			assertTrue(pos.toReference() + "の左に罫線がない事", !bordersInRange
					.getOnLeft().isAvailable());

		}

		if (!range.isTopEdge(pos)) {
			assertTrue(pos.toReference() + "の上に罫線がない事", !bordersInRange
					.getOnTop().isAvailable());
		}

		if (!range.isRightEdge(pos)) {
			assertTrue(pos.toReference() + "の右に罫線がない事", !bordersInRange
					.getOnRight().isAvailable());
		}

		if (!range.isBottomEdge(pos)) {
			assertTrue(pos.toReference() + "の下に罫線がない事", !bordersInRange
					.getOnBottom().isAvailable());
		}

	}

	private void assertBottomBorderOnlyIfEdge(ScPos pos,
			CellBorders bordersInRange, Border expectBorder, In range) {
		if (range.isBottomEdge(pos)) {
			assertBottomBorder(pos, bordersInRange, expectBorder);

		} else {
			assertTrue(pos.toReference() + "の下に罫線がない事", !bordersInRange
					.getOnBottom().isAvailable());
		}
	}

	private void assertTopBorderOnlyIfEdge(ScPos pos,
			CellBorders bordersInRange, Border expectBorder, In range) {
		if (range.isTopEdge(pos)) {
			assertTopBorder(pos, bordersInRange, expectBorder);
		} else {
			assertTrue(pos.toReference() + "の上に罫線が無い事", !bordersInRange
					.getOnTop().isAvailable());
		}
	}

	private void assertRightBorderOnlyIfEdge(ScPos pos,
			CellBorders bordersInRange, Border expectBorder, In range) {
		if (range.isRightEdge(pos)) {
			assertRightBorder(pos, bordersInRange, expectBorder);
		} else {
			assertTrue(pos.toReference() + "の右に罫線が無い事", !bordersInRange
					.getOnRight().isAvailable());
		}
	}

	private void assertLeftBorderOnlyIfEdge(ScPos pos,
			CellBorders bordersInRange, Border expectBorder, In range) {
		if (range.isLeftEdge(pos)) {
			assertLeftBorder(pos, bordersInRange, expectBorder);
		} else {
			assertTrue(pos.toReference() + "の左に罫線が無い事", !bordersInRange
					.getOnLeft().isAvailable());
		}
	}

	public void isSurroundInBoxWith(final Border border, final In range)
			throws IOException {
		isSurroundInBoxWith(border, border, border, border, range);
	}

	public void isSurroundInBoxWith(final Border left, final Border top,
			final Border right, final Border bottom, final In range)
			throws IOException {

		CellBorders e = CellBorders.bordersToCellBorder(left, top, right,
				bottom);

		isSurroundInBoxWith(e, range);
	}

	public void isSurroundInBoxWith(final CellBorders expectBorders,
			final In range) throws IOException {

		new Assert(sheet, range) {

			@Override
			public void withBorders(ScCell cell, CellBorders bordersOfTheCell,
					BordersInSheet borders) {
				assertAroundBoxIfApropriateCell(cell.getPos(),
						bordersOfTheCell, expectBorders, range);
			}

		}.doAssert();

	}

	// TODO UT
	public void noInternalBorderIn(final In range) throws IOException {

		new Assert(sheet, range) {

			@Override
			public void withBorders(ScCell cell, CellBorders bordersOfTheCell,
					BordersInSheet borders) {

				assertNoInternalLineIfApropriateCell(cell.getPos(),
						bordersOfTheCell, range);
			}

		}.doAssert();

	}

	public void isLeftBorder(final Border expect, final In range)
			throws FileNotFoundException, IOException {
		new Assert(sheet, range) {
			@Override
			public void withBorders(ScCell cell, CellBorders bordersOfTheCell,
					BordersInSheet borders) {

				ScPos pos = cell.getPos();

				if (range.isLeftEdge(pos)) {
					assertLeftBorder(pos, bordersOfTheCell, expect);
				}
			}
		}.doAssert();
	}

	public void isRightBorder(final Border expect, final In range)
			throws FileNotFoundException, IOException {
		new Assert(sheet, range) {
			@Override
			public void withBorders(ScCell cell, CellBorders bordersOfTheCell,
					BordersInSheet borders) {

				ScPos pos = cell.getPos();

				if (range.isRightEdge(pos)) {
					assertRightBorder(pos, bordersOfTheCell, expect);
				}
			}
		}.doAssert();
	}

	public void isTopBorder(final Border expect, final In range)
			throws FileNotFoundException, IOException {
		new Assert(sheet, range) {
			@Override
			public void withBorders(ScCell cell, CellBorders bordersOfTheCell,
					BordersInSheet borders) {

				ScPos pos = cell.getPos();

				if (range.isTopEdge(pos)) {
					assertTopBorder(pos, bordersOfTheCell, expect);
				}
			}
		}.doAssert();
	}

	public void isBottomBorder(final Border expect, final In range)
			throws FileNotFoundException, IOException {
		new Assert(sheet, range) {
			@Override
			public void withBorders(ScCell cell, CellBorders bordersOfTheCell,
					BordersInSheet borders) {

				ScPos pos = cell.getPos();

				if (range.isBottomEdge(pos)) {
					assertBottomBorder(pos, bordersOfTheCell, expect);
				}
			}
		}.doAssert();
	}

	public static abstract class Assert {

		private final ScSheet sheet;
		private final In range;

		public Assert(ScSheet sheet, In range) {
			this.sheet = sheet;
			this.range = range;
		}

		public abstract void withBorders(ScCell cell,
				CellBorders bordersOfTheCell, BordersInSheet bordersInRange)
				throws FileNotFoundException, IOException;

		public void doAssert() throws FileNotFoundException, IOException {

			In fetchBorderRange = createBorderCheckRange(range);

			final BordersInSheet borders = BordersInSheet.fetchFromSheet(
					new ScPos(fetchBorderRange.start()), new ScPos(
							fetchBorderRange.end()), sheet

			);

			new AssertIn(sheet, range) {
				@Override
				public void that(ScCell cell) throws FileNotFoundException,
						IOException {

					ScPos posOfTheCell = cell.getPos();
					CellBorders bordersOfTheCell = borders.getMatrix().get(
							posOfTheCell);

					assertTrue(posOfTheCell.toReference() + "がnullではない",
							bordersOfTheCell != null);

					withBorders(cell, bordersOfTheCell, borders);
				}
			}.doAssert();

		}

		private In createBorderCheckRange(final In range) {
			// TODO

			final String start = new ScPos(range.start()).moveInRow(-1)
					.moveInCol(-1).toReference();
			final String end = new ScPos(range.end()).moveInCol(1).moveInRow(1)
					.toReference();

			In fetchBorderRange = new In(start, end) {

				@Override
				public String end() {
					return end;
				}

				@Override
				public String start() {
					return start;
				}
			};
			return fetchBorderRange;
		}
	}

}
