package org.zkoss.zss.api;

import org.zkoss.zss.api.impl.RangeImpl;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;

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
	
	/** Returns the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4" or "Sheet2!A1:D4") or name of a NamedRange (e.g. "MyRange");
	 * note that if reference string contains sheet name, the returned range will refer to the named sheet. 
	 *  
	 * @param sheet the {@link Sheet} the Range will refer to.
	 * @param reference the area the Range will refer to (e.g. "A1:D4").
	 * @return the associated {@link Range} of the specified {@link Sheet} and area reference string (e.g. "A1:D4"). 
	 */
	public static Range range(Sheet sheet, String areaReference){
		return new RangeImpl(sheet,XRanges.range(((SheetImpl)sheet).getNative(),areaReference));
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
}
