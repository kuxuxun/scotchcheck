package com.github.kuxuxun.scotchtest.assertion.conf;

import java.io.File;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.ScWorkbook;
import com.github.kuxuxun.scotch.excel.ScWorkbookFactory;

public abstract class ReportFile {

	private static final Properties p = new Properties();

	protected abstract String getDir(Properties p);

	public File getFile(String fileName) {
		String dir = "";
		File f = null;
		try {
			p.load(this.getClass().getClassLoader().getResourceAsStream(
					"scotch.propertis"));
			f = new File(dir + fileName);
		} catch (IOException e) {
			throw new IllegalStateException("ファイルの読み込みに失敗しました:" + dir
					+ fileName);

		}

		return f;

	}

	public ScWorkbook getWorkbook(String fileName) {
		ScWorkbook wb;
		String dir = "";
		try {
			p.load(this.getClass().getClassLoader().getResourceAsStream(
					"scotch.propertis"));

			dir = getDir(p);
			wb = ScWorkbookFactory.createFromFile(new File(dir + fileName));

		} catch (IOException e) {
			throw new IllegalStateException("ワークブックの読み込みに失敗しました:" + dir
					+ fileName);
		} catch (InvalidFormatException e) {
			throw new IllegalStateException("ワークブックの読み込みに失敗しました:" + dir
					+ fileName);
		}

		return wb;
	}

	public static <T extends ReportFile> ScSheet getFirstSheetOf(
			String fileName, Class<T> c) {
		ScWorkbook wb;
		try {
			wb = workbook(fileName, c);
		} catch (Exception ex) {
			throw new IllegalArgumentException("シートの取得に失敗しました。 " + fileName
					+ " msg: " + ex.getMessage());
		}

		return wb.getFirstSheet();
	}

	public static <T extends ReportFile> ScWorkbook workbook(String fileName,
			Class<T> c) {
		try {
			return (c.newInstance()).getWorkbook(fileName);

		} catch (InstantiationException e) {
			throw new IllegalArgumentException("ファイル読み込みに失敗しました" + fileName);
		} catch (IllegalAccessException e) {
			throw new IllegalArgumentException("ファイル読み込みに失敗しました" + fileName);
		}
	}

	public static <T extends ReportFile> String dirPathOf(Class<T> c) {
		try {
			return (c.newInstance()).getDir(p);

		} catch (InstantiationException e) {
			throw new IllegalArgumentException("ファイル読み込みに失敗しました");
		} catch (IllegalAccessException e) {
			throw new IllegalArgumentException("ファイル読み込みに失敗しました");
		}
	}

	public static <T extends ReportFile> File file(String fileName, Class<T> c) {
		try {
			return (c.newInstance()).getFile(fileName);

		} catch (InstantiationException e) {
			throw new IllegalArgumentException("ファイル読み込みに失敗しました" + fileName);
		} catch (IllegalAccessException e) {
			throw new IllegalArgumentException("ファイル読み込みに失敗しました" + fileName);
		}
	}

}
