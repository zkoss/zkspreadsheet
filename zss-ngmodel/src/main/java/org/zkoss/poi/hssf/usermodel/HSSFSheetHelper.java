/* HSSFSheetHelper.java

	Purpose:
		
	Description:
		
	History:
		Jun 4, 2010 10:28:41 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.poi.hssf.usermodel;

import org.zkoss.poi.hssf.model.InternalSheet;
import org.zkoss.poi.hssf.usermodel.HSSFSheet;

/**
 * Copied from ZSS project temporarily to make BookHelper work.
 */
public class HSSFSheetHelper {
	private final HSSFSheet _sheet;
	public HSSFSheetHelper(HSSFSheet sheet) {
		_sheet = sheet;
	}
	public HSSFSheet getSheet() {
		return _sheet;
	}
	public InternalSheet getInternalSheet() {
		return _sheet.getSheet();
	}
}
