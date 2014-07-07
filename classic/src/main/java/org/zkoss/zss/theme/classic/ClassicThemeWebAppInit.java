package org.zkoss.zss.theme.classic;

import org.zkoss.zk.ui.WebApp;
import org.zkoss.zk.ui.util.WebAppInit;
import org.zkoss.zss.theme.SpreadsheetThemes;

public class ClassicThemeWebAppInit implements WebAppInit {

	private final static String CLASSIC_NAME = "classic";
	private final static String CLASSIC_DISPLAY = "Classic";
	private final static int CLASSIC_PRIORITY = 1000;

	public void init(WebApp wapp) throws Exception {
		SpreadsheetThemes.register(CLASSIC_NAME, CLASSIC_DISPLAY, CLASSIC_PRIORITY);
	}
}
