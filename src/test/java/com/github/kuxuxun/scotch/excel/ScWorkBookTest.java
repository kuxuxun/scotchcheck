package com.github.kuxuxun.scotch.excel;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.github.kuxuxun.commonutil.ut.TestExpected;
import com.github.kuxuxun.commonutil.ut.TestOutput;
import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.ScWorkbook;
import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.cell.ScCell;
import com.github.kuxuxun.scotch.excel.cell.ScFormulas;

public class ScWorkBookTest {

	@Test
	public void assertFixCellValue() throws Exception {

		ScWorkbook expected = TestExpected.fileOf("expected.xls");

		ScSheet e = expected.getSheetAt(0);

		ScWorkbook wb = new ScWorkbook(new HSSFWorkbook());

		ScSheet t = wb.createSheet("test");

		ScCell c3 = e.getCellAt("C3");
		ScCell c4 = e.getCellAt("C4");

		t.getCellAt("C3").setTextWithStyle("fuck", c3.getStyle());
		t.getCellAt("C4").setTextWithStyle("fuck", c4.getStyle());

		t.getCellAt("C3")
				.setFunctionWithStyle(
						ScFormulas.createRateFunction(new ScPos("A1"),
								new ScPos("A2")), c3.getStyle());

		Workbook report = wb.getPoiWorkBook();
		FileOutputStream fileOut = TestOutput.getAsOutputStream("output.xls");
		report.write(fileOut);
		fileOut.close();

	}
}
