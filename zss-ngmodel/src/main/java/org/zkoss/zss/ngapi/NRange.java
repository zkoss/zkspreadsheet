package org.zkoss.zss.ngapi;

import java.util.Locale;

import org.zkoss.zss.ngmodel.NSheet;

public interface NRange {

	NSheet getSheet();
	int getRow();
	int getColumn();
	int getLastRow();
	int getLastColumn();
	
	void setEditText(String editText);
	void setValue(Object value);
	void clear();
	void setLocale(Locale locale);
	Locale getLocale();
}
