package org.zkoss.zss.api;

import org.zkoss.zss.api.impl.RangeImpl;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.sys.XRanges;

public class Ranges {

	public static Range range(Sheet sheet){
		return new RangeImpl(sheet,XRanges.range(((SheetImpl)sheet).getNative()));
	}
	
//	public static NRange range(NSheet sheet, String areaReference){
//		return new NRangeImpl(sheet,Ranges.range(((NSheetImpl)sheet).getNative(),areaReference));
//	}
	
	public static Range range(Sheet sheet, int tRow, int lCol, int bRow, int rCol){
		return new RangeImpl(sheet,XRanges.range(((SheetImpl)sheet).getNative(),tRow,lCol,bRow,rCol));
	}
	
	public static Range range(Sheet sheet, int row, int col){	
		return new RangeImpl(sheet,XRanges.range(((SheetImpl)sheet).getNative(),row,col));
	}
}
