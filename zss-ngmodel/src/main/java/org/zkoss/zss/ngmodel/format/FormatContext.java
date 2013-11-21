package org.zkoss.zss.ngmodel.format;

import java.util.Locale;

public class FormatContext {

	Locale locale;

	public FormatContext() {
		this.locale = Locale.getDefault();
	}

	public FormatContext(Locale locale) {
		this.locale = locale;
	}

	public Locale getLocale() {
		return locale;
	}

	public void setLocale(Locale locale) {
		this.locale = locale;
	}
}
