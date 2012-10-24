package com.github.kuxuxun.scotchtest.assertion;

import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestExpected;
import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotchtest.assertion.bdd.シート;

//FIXME compareTarget.xlsをexectedに寄せる
public class AssertAlignTest {
	@Test
	public void test左詰め_下詰め() throws Exception {

		ScSheet s = TestExpected.getFirstSheetOf("compareTarget.xls");
		new シート(s).のセル("B9").の文字列が("横：左詰め、縦：下詰め");

		new シート(s).のセル("B9").のスタイル().の横位置が左詰め();

		new シート(s).のセル("B9").のスタイル().の縦位置が下詰め();

	}

	@Test
	public void test右詰め_縦両端そろえ() throws Exception {

		ScSheet s = TestExpected.getFirstSheetOf("compareTarget.xls");
		new シート(s).のセル("B10").の文字列が("横右つめ、縦：両端そろえ");

		new シート(s).のセル("B10").のスタイル().の横位置が右詰め();

		new シート(s).のセル("B10").のスタイル().の縦位置が両端そろえ();

	}
}
