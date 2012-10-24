package com.github.kuxuxun.scotchtest.assertion.bdd;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotchtest.assertion.assertion.AssertStyle;
import com.github.kuxuxun.scotchtest.assertion.assertion.In;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.HAlign;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.VAlign;

public class スタイル {
	private final ScSheet targetSheet;
	private final In range;

	public スタイル(ScSheet targetSheet, In range) {
		this.targetSheet = targetSheet;
		this.range = range;
	}

	public void の横位置が両端揃え() throws FileNotFoundException, IOException {
		の横位置が(HAlign.JUSTIFY);
	}

	public void の横位置が左詰め() throws FileNotFoundException, IOException {
		の横位置が(HAlign.LEFT);
	}

	public void の横位置が右詰め() throws FileNotFoundException, IOException {
		の横位置が(HAlign.RIGHT);
	}

	public void の横位置が(HAlign ha) throws FileNotFoundException, IOException {
		AssertStyle.that(targetSheet).horizontalAlignIs(ha, range);
	}

	public void の縦位置が両端そろえ() throws FileNotFoundException, IOException {
		の縦位置が(VAlign.VERTICAL_JUSTIFY);
	}

	public void の縦位置が下詰め() throws FileNotFoundException, IOException {
		の縦位置が(VAlign.VERTICAL_BOTTOM);
	}

	public void の縦位置が(VAlign va) throws FileNotFoundException, IOException {
		AssertStyle.that(targetSheet).verticalAllignIs(va, range);
	}

}
