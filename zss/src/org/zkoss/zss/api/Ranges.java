/* Ranges.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api;

import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.zss.api.impl.RangeImpl;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.Rect;

/**
 * The facade class provides you multiple ways to get a {@link Range}.
 * 
 * @author dennis
 * @see Range 
 * @since 3.0.0
 */
public class Ranges {

	/** 
	 * Returns the associated {@link Range} of the whole specified {@link Sheet}. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @return the associated {@link Range} of the whole specified {@link Sheet}. 
	 */
	public static Range range(Sheet sheet){
		return new RangeImpl(sheet,XRanges.range(((SheetImpl)sheet).getNative()));
	}
	
	/** Returns the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4" or "Sheet2!A1:D4")
	 * note that if reference string contains sheet name, the returned range will refer to the named sheet. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param reference the area the Range will refer to (e.g. "A1:D4").
	 * @return the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4"). 
	 */
	public static Range range(Sheet sheet, String areaReference){
		return new RangeImpl(sheet,XRanges.range(((SheetImpl)sheet).getNative(),areaReference));
	}
	
	/** Returns the associated {@link Range} of the specified name of a NamedRange (e.g. "MyRange");
	 * 
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param name the name of NamedRange  (e.g. "MyRange"); .
	 * @return the associated {@link Range} of the specified name 
	 */
	public static Range rangeByName(Sheet sheet, String name){
		return new RangeImpl(sheet,XRanges.rangeByName(((SheetImpl)sheet).getNative(),name));
	}	
	
	/** Returns the associated {@link XRange} of the specified {@link XSheet} and area. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param tRow top row index of the area the Range will refer to.
	 * @param lCol left column index of the area the Range will refer to.
	 * @param bRow bottom row index of the area the Range will refer to.
	 * @param rCol right column index fo the area the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Sheet} and area.
	 */
	public static Range range(Sheet sheet, int tRow, int lCol, int bRow, int rCol){
		return new RangeImpl(sheet,XRanges.range(((SheetImpl)sheet).getNative(),tRow,lCol,bRow,rCol));
	}
	
	/** Returns the associated {@link Range} of the specified {@link Sheet} and cell row and column. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param row row index of the cell the Range will refer to.
	 * @param col column index of the cell the Range will refer to.
	 * @return the associated {@link Range} of the specified {@link Sheet} and cell . 
	 */
	public static Range range(Sheet sheet, int row, int col){	
		return new RangeImpl(sheet,XRanges.range(((SheetImpl)sheet).getNative(),row,col));
	}
	
	/** Returns the associated {@link Range} of the specified {@link Sheet} and cell row and column. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param selection the selection of spreadsheet
	 */
	public static Range range(Sheet sheet, Rect selection){	
		return range(sheet,selection.getTop(),selection.getLeft(),selection.getBottom(),selection.getRight());
	}
	
	/**
	 * Gets cell reference expression
	 * @param row row
	 * @param col column
	 * @return the cell reference string (e.g. A1)
	 */
	public static String toCellReference(int row,int col){
		return toCellReference(null,row,col);
	}
	
	/**
	 * Gets cell reference expression
	 * @param sheet sheet
	 * @param row row
	 * @param col column
	 * @return the cell reference string (e.g. Sheet1!A1)
	 */
	public static String toCellReference(Sheet sheet,int row,int col){
		CellReference cf = new CellReference(sheet==null?null:sheet.getSheetName(),row, col,false,false);
		return cf.formatAsString();
	}
	
	/**
	 * Gets area reference expression
	 * @param tRow top row
	 * @param lCol left column
	 * @param bRow bottom row
	 * @param rCol right column
	 * @return the area reference string (e.g. A1:B2)
	 */
	public static String toAreaReference(int tRow, int lCol, int bRow, int rCol){
		return toAreaReference(null,tRow, lCol, bRow, rCol);
	}
	
	/**
	 * Gets area reference expression
	 * @param area area
	 * @return the area reference string (e.g. A1:B2)
	 */
	public static String toAreaReference(Rect area){
		return toAreaReference(area.getTop(),area.getLeft(),area.getBottom(),area.getRight());
	}	
	
	/**
	 * Gets area reference expression
	 * @param sheet sheet
	 * @param tRow top row
	 * @param lCol left column
	 * @param bRow bottom row
	 * @param rCol right column
	 * @return the area reference string (e.g. Sheet1!A1:B2)
	 */
	public static String toAreaReference(Sheet sheet, int tRow, int lCol, int bRow, int rCol){
		String sn = sheet==null?null:sheet.getSheetName();
		AreaReference af = new AreaReference(new CellReference(sn,tRow,lCol,false,false), new CellReference(sn,bRow,rCol,false,false));
		return af.formatAsString();
	}
	
	/**
	 * Gets area reference expression
	 * @param sheet sheet
	 * @param area area
	 * @return the area reference string (e.g. Sheet1!A1:B2)
	 */
	public static String toAreaReference(Sheet sheet, Rect area){
		String sn = sheet==null?null:sheet.getSheetName();
		return toAreaReference(sheet, area.getTop(),area.getLeft(),area.getBottom(),area.getRight());
	}
	
	/**
	 * Get column reference
	 * @param column column
	 * @return the column reference string (e.g A, AB)
	 */
	public static String toColumnReference(int column){
		return CellReference.convertNumToColString(column);
	}
	
	/**
	 * Get row reference
	 * @param row row
	 * @return the column reference string (e.g 1, 12)
	 */
	public static String toRowReference(int row){
		int excelRowNum = row + 1;
		return Integer.toString(excelRowNum);
	}
}
