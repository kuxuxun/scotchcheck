package com.github.kuxuxun.scotchtest.assertion.conf;

import java.util.Properties;

public class Outputted extends ReportFile {

	@Override
	protected String getDir(Properties p) {
		return p.getProperty("output.dir");
	}

}
