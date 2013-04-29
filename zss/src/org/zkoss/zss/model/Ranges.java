/* Ranges.java

	Purpose:
		
	Description:
		
	History:
		Sep 6, 2010 10:56:23 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import org.zkoss.poi.ss.util.AreaReference;


/**
 * Internal Use Only. Utilities regarding Range operation.
 * @author henrichen
 * @author dennischen
 * @deprecated since 3.0.0, please use class in package {@code org.zkoss.zss.api}
 */
public class Ranges {
	/** Returns the associated {@link Range} of the whole specified {@link Worksheet}. 
	 *  
	 * @param sheet the {@link Worksheet} the Range will refer to.
	 * @return the associated {@link Range} of the whole specified {@link Worksheet}. 
	 */
	public static Range range(Worksheet sheet) {
		throw new ModelException("the api was not support anymore, please use api in org.zkoss.zss.api");
	}
	
	/** Returns the associated {@link Range} of the specified {@link Worksheet} and area reference string (e.g. "A1:D4" or "Sheet2!A1:D4") or name of a NamedRange (e.g. "MyRange");
	 * note that if reference string contains sheet name, the returned range will refer to the named sheet. 
	 *  
	 * @param sheet the {@link Worksheet} the Range will refer to.
	 * @param reference the area the Range will refer to (e.g. "A1:D4").
	 * @return the associated {@link Range} of the specified {@link Worksheet} and area reference string (e.g. "A1:D4"). 
	 */
	public static Range range(Worksheet sheet, String reference) {
		throw new ModelException("the api was not support anymore, please use api in org.zkoss.zss.api");
	}
	
	
	public static Range range(Worksheet sheet, AreaReference ref) {
		throw new ModelException("the api was not support anymore, please use api in org.zkoss.zss.api");
	}
	
	/** Returns the associated {@link Range} of the specified {@link Worksheet} and cell row and column. 
	 *  
	 * @param sheet the {@link Worksheet} the Range will refer to.
	 * @param row row index of the cell the Range will refer to.
	 * @param col column index of the cell the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Worksheet} and cell . 
	 */
	public static Range range(Worksheet sheet, int row, int col) {
		return newRange(sheet, sheet, row, col, row, col);
	}

	/** Returns the associated {@link Range} of the specified {@link Worksheet} and area. 
	 *  
	 * @param sheet the {@link Worksheet} the Range will refer to.
	 * @param tRow top row index of the area the Range will refer to.
	 * @param lCol left column index of the area the Range will refer to.
	 * @param bRow bottom row index of the area the Range will refer to.
	 * @param rCol right column index fo the area the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Worksheet} and area.
	 */
	public static Range range(Worksheet sheet, int tRow, int lCol, int bRow, int rCol) {
		return newRange(sheet, sheet, tRow, lCol, bRow, rCol);
	}
	
	/** Returns the associated three dimension {@link Range} of the specified {@link Worksheet}s 
	 * and cell row and column. 
	 *  
	 * @param firstSheet the first {@link Worksheet} the Range will start referring to.
	 * @param lastSheet the last {@link Worksheet} the Range will end referring to.
	 * @param row row index of the cell the Range will refer to.
	 * @param col column index of the cell the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Worksheet} and cell . 
	 */
	public static Range range(Worksheet firstSheet, Worksheet lastSheet, int row, int col) {
		return newRange(firstSheet, lastSheet, row, col, row, col);
	}
	
	/** Returns the associated three dimension {@link Range} of the specified {@link Worksheet}s and area. 
	 *  
	 * @param firstSheet the first {@link Worksheet} the Range will start referring to.
	 * @param lastSheet the last {@link Worksheet} the Range will end referring to.
	 * @param tRow top row index of the area the Range will refer to.
	 * @param lCol left column index of the area the Range will refer to.
	 * @param bRow bottom row index of the area the Range will refer to.
	 * @param rCol right column index fo the area the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Worksheet} and area.
	 */
	public static Range range(Worksheet firstSheet, Worksheet lastSheet, int tRow, int lCol, int bRow, int rCol) {
		return newRange(firstSheet, lastSheet, tRow, lCol, bRow, rCol);
	}
	
	private static Range newRange(Worksheet firstSheet, Worksheet lastSheet, int tRow, int lCol, int bRow, int rCol) {
		throw new ModelException("the api was not support anymore, please use api in org.zkoss.zss.api");
	}
}
