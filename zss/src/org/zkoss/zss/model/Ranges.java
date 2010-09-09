/* Ranges.java

	Purpose:
		
	Description:
		
	History:
		Sep 6, 2010 10:56:23 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.zkoss.lang.Objects;
import org.zkoss.zss.model.impl.RangeImpl;

/**
 * Utilities regarding Range operation.
 * @author henrichen
 *
 */
public class Ranges {
	/** Returns the associated {@link Range} of the whole specified {@link Sheet}. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @return the associated {@link Range} of the whole specified {@link Sheet}. 
	 */
	public static Range range(Sheet sheet) {
		final Book book = (Book) sheet.getWorkbook();
		final SpreadsheetVersion ver = book.getSpreadsheetVersion();
		return newRange(sheet, sheet, 0, 0, ver.getLastRowIndex(), ver.getLastColumnIndex());
	}
	
	/** Returns the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4"). 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param reference the area the Range will refer to (e.g. "A1:D4").
	 * @return the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4"). 
	 */
	public static Range range(Sheet sheet, String reference) {
		AreaReference ref = new AreaReference(reference);
		final int tRow = ref.getFirstCell().getRow();
		final int lCol = ref.getFirstCell().getCol();
		final int bRow = ref.getLastCell().getRow();
		final int rCol = ref.getLastCell().getCol();
		return newRange(sheet, sheet, tRow, lCol, bRow, rCol);
	}
	
	/** Returns the associated three dimension {@link Range} of the specified {@link Sheet}s and area reference string (e.g. "A1:D4"). 
	 * 
	 * @param firstSheet the first {@link Sheet} the Range will start referring to.
	 * @param lastSheet the last {@link Sheet} the Range will end referring to.
	 * @param reference the area the Range will refer to (e.g. "A1:D4").
	 * @return the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4"). 
	 */
	public static Range range(Sheet firstSheet, Sheet lastSheet, String reference) {
		AreaReference ref = new AreaReference(reference);
		final int tRow = ref.getFirstCell().getRow();
		final int lCol = ref.getFirstCell().getCol();
		final int bRow = ref.getLastCell().getRow();
		final int rCol = ref.getLastCell().getCol();
		
		return newRange(firstSheet, lastSheet, tRow, lCol, bRow, rCol);
	}
	
	/** Returns the associated {@link Range} of the specified {@link Sheet} and cell row and column. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param row row index of the cell the Range will refer to.
	 * @param col column index of the cell the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Sheet} and cell . 
	 */
	public static Range range(Sheet sheet, int row, int col) {
		return newRange(sheet, sheet, row, col, row, col);
	}

	/** Returns the associated {@link Range} of the specified {@link Sheet} and area. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param tRow top row index of the area the Range will refer to.
	 * @param lCol left column index of the area the Range will refer to.
	 * @param bRow bottom row index of the area the Range will refer to.
	 * @param rCol right column index fo the area the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Sheet} and area.
	 */
	public static Range range(Sheet sheet, int tRow, int lCol, int bRow, int rCol) {
		return newRange(sheet, sheet, tRow, lCol, bRow, rCol);
	}
	
	/** Returns the associated three dimension {@link Range} of the specified {@link Sheet}s 
	 * and cell row and column. 
	 *  
	 * @param firstSheet the first {@link Sheet} the Range will start referring to.
	 * @param lastSheet the last {@link Sheet} the Range will end referring to.
	 * @param row row index of the cell the Range will refer to.
	 * @param col column index of the cell the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Sheet} and cell . 
	 */
	public static Range range(Sheet firstSheet, Sheet lastSheet, int row, int col) {
		return newRange(firstSheet, lastSheet, row, col, row, col);
	}
	
	/** Returns the associated three dimension {@link Range} of the specified {@link Sheet}s and area. 
	 *  
	 * @param firstSheet the first {@link Sheet} the Range will start referring to.
	 * @param lastSheet the last {@link Sheet} the Range will end referring to.
	 * @param tRow top row index of the area the Range will refer to.
	 * @param lCol left column index of the area the Range will refer to.
	 * @param bRow bottom row index of the area the Range will refer to.
	 * @param rCol right column index fo the area the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Sheet} and area.
	 */
	public static Range range(Sheet firstSheet, Sheet lastSheet, int tRow, int lCol, int bRow, int rCol) {
		return newRange(firstSheet, lastSheet, tRow, lCol, bRow, rCol);
	}
	
	private static Range newRange(Sheet firstSheet, Sheet lastSheet, int tRow, int lCol, int bRow, int rCol) {
		return tRow == bRow && lCol == rCol ? 
				new RangeImpl(tRow, lCol, firstSheet, lastSheet):
				new RangeImpl(tRow, lCol, bRow, rCol, firstSheet, lastSheet);
	}
}
