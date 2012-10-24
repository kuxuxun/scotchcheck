package com.github.kuxuxun.scotchtest.assertion;

import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestExpected;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertTwoReport;

public class ReportComparatorTest {

	@Test
	public void equalBorderTest() throws Exception {
		new AssertTwoReport(

		TestExpected.getFirstSheetOf("compareTarget.xls")

		).isEqualsInValueAndBorderTo("expected.xls").in("D4", "G7");
	}

}
