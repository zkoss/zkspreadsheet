package org.zkoss.zss.ngapi;

import org.zkoss.zss.ngmodel.NSheet;

public interface NRange {

	NSheet getSheet();
	int getRow();
	int getColumn();
	int getLastRow();
	int getLastColumn();
	
	void setEditText(String editText);
}
