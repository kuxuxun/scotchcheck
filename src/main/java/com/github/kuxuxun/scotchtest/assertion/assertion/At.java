package com.github.kuxuxun.scotchtest.assertion.assertion;

import com.github.kuxuxun.scotch.excel.area.ScPos;

public class At extends In {

	public static At at(ScPos pos) {
		return At.at(pos.toReference());
	}

	public static At at(String at) {
		return new At(at);
	}

	private At(String at) {
		super(at, at);
	}

}