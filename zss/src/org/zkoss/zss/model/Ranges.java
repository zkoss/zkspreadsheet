/* Ranges.java

	Purpose:
		
	Description:
		
	History:
		Sep 6, 2010 10:56:23 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import java.util.Collection;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.zkoss.lang.Objects;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.EmptyRange;
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
	
	/** Returns the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4" or "Sheet2!A1:D4");
	 * note that if reference string contains sheet name, the returned range will refer to the named sheet. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param reference the area the Range will refer to (e.g. "A1:D4").
	 * @return the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4"). 
	 */
	public static Range range(Sheet sheet, String reference) {
		AreaReference ref = new AreaReference(reference);
		return range(sheet, ref);
	}
	
	/** Returns the associated {@link Range} of the specified {@link Sheet} and AreaReference. note that if AreaReference 
	 * contains sheet name, the returned range will refer to the named sheet. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param ref the AreaReference the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4"). 
	 */
	public static Range range(Sheet sheet, AreaReference ref) {
		final CellReference c1 = ref.getFirstCell();
		final CellReference c2 = ref.getLastCell();
		final String sheetnm1 = c1.getSheetName();
		final String sheetnm2 = c2.getSheetName();
		final Workbook book = sheet.getWorkbook(); 
		final Sheet sheet1 = sheetnm1 != null ? book.getSheet(sheetnm1) : sheet;
		final Sheet sheet2 = sheetnm2 != null ? book.getSheet(sheetnm2) : sheet;
		final int tRow = c1.getRow();
		final int lCol = c1.getCol();
		final int bRow = c2.getRow();
		final int rCol = c2.getCol();
		return newRange(sheet1, sheet2, tRow, lCol, bRow, rCol);
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
	
	public static final Range EMPTY_RANGE = new EmptyRange();
}
