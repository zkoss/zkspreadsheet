package org.zkoss.zss.ngapi;

import org.zkoss.zss.ngapi.impl.NRangeImpl;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.AreaReference;
import org.zkoss.zss.ngmodel.util.CellReference;


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
		AreaReference ar = new AreaReference(areaReference);
		CellReference cr1 = ar.getFirstCell();
		CellReference cr2 = ar.getLastCell();
		return new NRangeImpl(sheet,cr1.getRow(),cr1.getCol(),cr2.getRow(),cr2.getCol());
	}
	
//	/** Returns the associated {@link Range} of the specified name of a NamedRange (e.g. "MyRange");
//	 * 
//	 * @param sheet the {@link NSheet} the Range will refer to.
//	 * @param name the name of NamedRange  (e.g. "MyRange"); .
//	 * @return the associated {@link Range} of the specified name 
//	 */
//	public static Range rangeByName(NSheet sheet, String name){
//		return new RangeImpl(XRanges.rangeByName(((NSheetImpl)sheet).getNative(),name),sheet);
//	}	
	
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
