package com.github.kuxuxun.scotchtest.assertion.bdd;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertMergedCells;
import com.github.kuxuxun.scotchtest.assertion.assertion.In;

public class 範囲チェック {
	private final ScSheet targetSheet;
	private final In range;

	public 範囲チェック(ScSheet targetSheet, In range) {
		this.targetSheet = targetSheet;
		this.range = range;
	}

	public void が結合されている() throws FileNotFoundException, IOException {
		AssertMergedCells.that(targetSheet).cellsAreMerged(range);
	}

	public void が結合されていない() throws FileNotFoundException, IOException {
		AssertMergedCells.that(targetSheet).cellsAreNotMerged(range);
	}
}
