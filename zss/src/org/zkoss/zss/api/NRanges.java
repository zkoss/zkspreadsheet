package org.zkoss.zss.api;

import org.zkoss.zss.api.impl.NRangeImpl;
import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.api.model.impl.NSheetImpl;
import org.zkoss.zss.model.sys.Ranges;

public class NRanges {

	public static NRange range(NSheet sheet){
		return new NRangeImpl(sheet,Ranges.range(((NSheetImpl)sheet).getNative()));
	}
	
//	public static NRange range(NSheet sheet, String areaReference){
//		return new NRangeImpl(sheet,Ranges.range(((NSheetImpl)sheet).getNative(),areaReference));
//	}
	
	public static NRange range(NSheet sheet, int tRow, int lCol, int bRow, int rCol){
		return new NRangeImpl(sheet,Ranges.range(((NSheetImpl)sheet).getNative(),tRow,lCol,bRow,rCol));
	}
	
	public static NRange range(NSheet sheet, int row, int col){	
		return new NRangeImpl(sheet,Ranges.range(((NSheetImpl)sheet).getNative(),row,col));
	}
}
