/* Ranges.java

	Purpose:
		
	Description:
		
	History:
		Sep 6, 2010 10:56:23 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
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
	
	/** Returns the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4" or "Sheet2!A1:D4") or name of a NamedRange (e.g. "MyRange");
	 * note that if reference string contains sheet name, the returned range will refer to the named sheet. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param reference the area the Range will refer to (e.g. "A1:D4").
	 * @return the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4"). 
	 */
	public static Range range(Sheet sheet, String reference) {
		AreaReference ref = getAreaReference(sheet, reference);
		if (ref == null) {
			//try NamedRange
			final Workbook wb = sheet.getWorkbook();
		    final Name range = wb.getName(reference);
		    if (range != null) {
		    	ref = getAreaReference(sheet, range.getRefersToFormula());
		    } else {
		    	throw new IllegalArgumentException("Cannot find the named range '" + reference + "'");
		    }
		}
		if (ref == null)
			throw new IllegalArgumentException("Bad area reference '" + reference + "'");
		return range(sheet, ref);
	}
	
	private static AreaReference getAreaReference(Sheet sheet, String reference) {
		String[] parts = separateReference(reference);
		final String sheet1 = parts[0];
		final String sheet2 = parts[1];
		final String lt = parts[2];
		final String rb = parts[3];
		final String c1str = (sheet1 != null ? sheet1 : "") + (sheet2 != null ? (":" + sheet2) : "") + (sheet1 != null ? "!" : "") + lt;
		final SpreadsheetVersion ver = ((Book)sheet.getWorkbook()).getSpreadsheetVersion();  
		final int maxcol = ver.getLastColumnIndex();
		final int maxrow = ver.getLastRowIndex();

		AreaReference ref = null;
		try {
			final CellReference c1 = new CellReference(c1str);
			if (c1.getCol() <= maxcol && c1.getRow() <= maxrow) {
				final CellReference c2 = rb != null ? new CellReference(rb) : c1;
				if (rb != null) {
					if (c2.getCol() <= maxcol && c2.getRow() <= maxrow) {
						ref = new AreaReference(c1, c2);
					} 
				} else {
					ref = new AreaReference(c1, c2);
				}
			}
		} catch(Exception ex) {
			//ignore, return null ref and let upper case do it!
		}
		return ref;
	}
	
	//4 parts: sheet1, sheet2, left-top, right-bottom
	//null: not a legal reference string
	private static String[] separateReference(String reference) {
		StringBuffer sb = newSB();
		String sheet1 = null;
		String sheet2 = null;
		String lt = null;
		String rb = null;
		boolean quote = false;
		for(int j = 0, len = reference.length(); j < len; ++j) {
			final char ch = reference.charAt(j);
			switch(ch) {
			case '!': //Sheet delimiter
				if (!quote) { //end of sheet parts
					if (lt != null) {
						sheet1 = lt;
						lt = null;
						sheet2 = sb.toString();
					} else {
						sheet1 = sb.toString();
					}
					sb = newSB();
				} else {
					sb.append(ch); //in quote, so simple copy
				}
				break;
			case ':': //lt, rb delimiter; or sheet1, sheet2 delimiter
				lt = sb.toString();
				sb = newSB();
				break;
			case '\'': //single quote for sheet name
				quote = !quote;
				//fall under
			default:
				sb.append(ch);
			}
		}
		if (lt != null) {
			rb = sb.toString();
		} else {
			lt = sb.toString();
		}
		return new String[] {sheet1, sheet2, lt, rb};
	}
	
	private static StringBuffer newSB() {
		return new StringBuffer(32);
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
