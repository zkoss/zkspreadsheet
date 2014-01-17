/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi;

import org.zkoss.zss.ngapi.impl.NRangeImpl;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.InvalidateModelOpException;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NName;
import org.zkoss.zss.ngmodel.NSheet;

/**
 * To get the range.
 * @author dennis
 * @since 3.5.0
 */
public class NRanges {
	/** 
	 * Returns the associated {@link NRange} of the whole specified {@link NSheet}. 
	 *  
	 * @param sheet the {@link NSheet} the Range will refer to.
	 * @return the associated {@link NRange} of the whole specified {@link NSheet}. 
	 */
	public static NRange range(NSheet sheet){
		return new NRangeImpl(sheet);
	}
	
	/** Returns the associated {@link NRange} of the specified {@link NSheet} and area reference string (e.g. "A1:D4" or "NSheet2!A1:D4")
	 * note that if reference string contains sheet name, the returned range will refer to the named sheet. 
	 *  
	 * @param sheet the {@link NSheet} the Range will refer to.
	 * @param reference the area the Range will refer to (e.g. "A1:D4").
	 * @return the associated {@link NRange} of the specified {@link NSheet} and area reference string (e.g. "A1:D4"). 
	 */
	public static NRange range(NSheet sheet, String areaReference){
		CellRegion ar = new CellRegion(areaReference);
		return new NRangeImpl(sheet,ar.getRow(),ar.getColumn(),ar.getLastRow(),ar.getLastColumn());
	}
	
	/** Returns the associated {@link NRange} of the specified name of a NamedRange (e.g. "MyRange");
	 * 
	 * @param sheet the {@link NSheet} the Range will refer to.
	 * @param name the name of NamedRange  (e.g. "MyRange"); .
	 * @return the associated {@link NRange} of the specified name 
	 */
	public static NRange rangeByName(NSheet sheet, String name){
		NBook book = sheet.getBook();
		NName n = book.getNameByName(name);
		if(n==null){
			throw new InvalidateModelOpException("can't find name "+name);
		}
		sheet = book.getSheetByName(n.getRefersToSheetName());
		CellRegion region = n.getRefersToCellRegion();
		if(sheet==null || region==null){
			throw new InvalidateModelOpException("bad name "+name+ " : "+n.getRefersToFormula());
		}
		
		return new NRangeImpl(sheet,region.row,region.column,region.lastRow,region.lastColumn);
	}	
	
	/** Returns the associated {@link XRange} of the specified {@link XNSheet} and area. 
	 *  
	 * @param sheet the {@link NSheet} the Range will refer to.
	 * @param tRow top row index of the area the Range will refer to.
	 * @param lCol left column index of the area the Range will refer to.
	 * @param bRow bottom row index of the area the Range will refer to.
	 * @param rCol right column index fo the area the Range will refer to.
	 * @return the associated {@link NRange} of the specified {@link NSheet} and area.
	 */
	public static NRange range(NSheet sheet, int tRow, int lCol, int bRow, int rCol){
		return new NRangeImpl(sheet,tRow,lCol,bRow,rCol);
	}
	
	/** Returns the associated {@link NRange} of the specified {@link NSheet} and cell row and column. 
	 *  
	 * @param sheet the {@link NSheet} the Range will refer to.
	 * @param row row index of the cell the Range will refer to.
	 * @param col column index of the cell the Range will refer to.
	 * @return the associated {@link NRange} of the specified {@link NSheet} and cell . 
	 */
	public static NRange range(NSheet sheet, int row, int col){	
		return new NRangeImpl(sheet,row,col);
	}
}
