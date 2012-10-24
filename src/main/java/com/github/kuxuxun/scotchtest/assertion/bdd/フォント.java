package com.github.kuxuxun.scotchtest.assertion.bdd;

import com.github.kuxuxun.scotchtest.assertion.assertion.utils.FontExample;

//TODO インデント対応
public class フォント {

	public static FontExample 太さ通常_Alialサイズは(int size) {
		return FontExample.Builder.simpleArial(size).setBold(false).build();
	}

	public static FontExample 太いAlialサイズは(int size) {
		return FontExample.Builder.simpleArial(size).setBold(true).build();
	}

	public static FontExample 太文字斜体サイズは(int size) {
		return FontExample.Builder.simpleArial(size).setBold(true).setItalic(
				true).build();
	}

}
