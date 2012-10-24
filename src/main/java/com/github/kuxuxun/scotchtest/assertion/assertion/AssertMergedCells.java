package com.github.kuxuxun.scotchtest.assertion.assertion;

import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.area.ScRange;
import com.github.kuxuxun.scotch.excel.cell.ScCell;

public class AssertMergedCells {

	private final ScSheet sheet;

	private AssertMergedCells(ScSheet report) {
		this.sheet = report;
	}

	public static AssertMergedCells that(ScSheet file) {
		return new AssertMergedCells(file);
	}

	public void cellsAreMerged(In rangeCheckTo) throws FileNotFoundException,
			IOException {
		final List<ScRange> allMargedRangesInSheet = sheet.getMergedRanges();

		final ScRange rangeThatContainsTopleftCell = ScRange
				.selectWhichContainOrNull(rangeCheckTo.getPositions().get(0),
						allMargedRangesInSheet);

		if (rangeThatContainsTopleftCell == null) {
			throw new IllegalArgumentException("指定のセルを含む結合セルがありません："
					+ rangeCheckTo.getPositions().get(0).toString());
		}

		// チェック範囲の各セルが同一の結合範囲に含まれるか確認
		new AssertIn(sheet, rangeCheckTo) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				ScRange rangeThatContainsTheCell = ScRange
						.selectWhichContainOrNull(cell.getPos(),
								allMargedRangesInSheet);

				if (rangeThatContainsTheCell == null) {
					throw new IllegalArgumentException("指定のセルを含む結合セルがありません："
							+ cell.getPos().toString());
				}
				if (!rangeThatContainsTheCell
						.equals(rangeThatContainsTopleftCell)) {
					throw new IllegalArgumentException(cell.getPos()
							+ "は異なる結合範囲に含まれています。");
				}

			}
		}.doAssert();

		// チェック対象の範囲 = 結合セルの範囲であるか確認

		assertTrue("もっと広い範囲が結合されています。 指定:["
				+ rangeCheckTo.toScRange().toString() + "] ,but actual ["
				+ rangeThatContainsTopleftCell.toString() + "]",
				rangeThatContainsTopleftCell.equals(rangeCheckTo.toScRange()));

	}

	public void cellsAreNotMerged(In rangeCheckTo)
			throws FileNotFoundException, IOException {
		final List<ScRange> ranges = sheet.getMergedRanges();

		new AssertIn(sheet, rangeCheckTo) {
			@Override
			public void that(ScCell cell) throws FileNotFoundException,
					IOException {

				ScRange containingRange = ScRange.selectWhichContainOrNull(cell
						.getPos(), ranges);

				assertNull(
						cell.getPos().toReference() + "がマージ範囲に含まれていてはいけない: ",
						containingRange);
			}
		}.doAssert();

	}

}
