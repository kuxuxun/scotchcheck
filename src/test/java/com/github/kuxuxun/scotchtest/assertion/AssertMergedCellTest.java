package com.github.kuxuxun.scotchtest.assertion;

import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestExpected;
import com.github.kuxuxun.scotchtest.assertion.bdd.シート;

public class AssertMergedCellTest {
	@Test
	public void testマージセルテスト() throws Exception {
		new シート(TestExpected.getFirstSheetOf("compareTarget.xls")).のセル("B18")
				.から("C19").の範囲が().が結合されている();

		new シート(TestExpected.getFirstSheetOf("compareTarget.xls")).のセル("B16")
				.から("C17").の範囲が().が結合されていない();

	}

}
