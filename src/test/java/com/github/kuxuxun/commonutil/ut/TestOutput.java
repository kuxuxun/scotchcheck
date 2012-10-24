package com.github.kuxuxun.commonutil.ut;

import java.io.IOException;
import java.io.FileOutputStream;
import java.io.FileNotFoundException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.ScWorkbook;

public class TestOutput {
   

	public static FileOutputStream getAsOutputStream(String fileName) {

	    FileOutputStream fos = null;
	    try {
		return new FileOutputStream("/Temp/" + fileName);
	    } catch (FileNotFoundException ex) {
		new IllegalArgumentException("/Temp Not found");
	    }

	    return fos;

	}


}
