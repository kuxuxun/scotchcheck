package com.github.kuxuxun.scotchtest.assertion.assertion;

import static org.junit.Assert.assertTrue;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.cell.ScCell;

public abstract class AssertIn {
	private final In range;
	private final ScSheet sheet;

	public AssertIn(ScSheet sheet, In range) throws FileNotFoundException,
			IOException {
		this.range = range;
		this.sheet = sheet;
	}

	public abstract void that(ScCell cell) throws FileNotFoundException,
			IOException;

	public void doAssert() throws FileNotFoundException, IOException {

		List<ScPos> poses = range.getPositions();
		boolean assertAtLeaseOnce = false;

		for (ScPos each : poses) {
			ScCell cell = sheet.getCellAt(each);
			that(cell);
			assertAtLeaseOnce = true;
		}

		assertTrue("アサーションが最低一回は走っている", assertAtLeaseOnce);

	}
}