package org.zkoss.zss.api.impl;

import org.zkoss.zss.api.Range;

public class Util {
	
	public static boolean isAMergedRange(Range range) {
		org.zkoss.poi.ss.usermodel.Sheet sheet = range.getSheet().getPoiSheet();
		// go thorugh all region
		for(int number = sheet.getNumMergedRegions(); number > 0; number--) {
			org.zkoss.poi.ss.util.CellRangeAddress addr = sheet.getMergedRegion(number-1);
			// match four corner
			if(addr.getFirstRow() == range.getRow() && addr.getLastRow() == range.getLastRow()
					&& addr.getFirstColumn() == range.getColumn() && addr.getLastColumn() == range.getLastColumn()) {
				return true;
			}
		}
		return false; // doesn't match any region
	}

}
