package org.zkoss.zss.ngmodel.sys;

import java.util.Locale;

public class AbstractContext {

	Locale locale;

	public AbstractContext() {
		this.locale = Locale.getDefault();
	}

	public AbstractContext(Locale locale) {
		this.locale = locale;
	}

	public Locale getLocale() {
		return locale;
	}

	public void setLocale(Locale locale) {
		this.locale = locale;
	}
}
