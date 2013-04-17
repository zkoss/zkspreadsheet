package org.zkoss.zss.api;

import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.model.Ranges;

public class NRanges {

	public static NRange range(NSheet sheet){
		return new NRange(sheet,Ranges.range(sheet.getNative()));
	}
	
	public static NRange range(NSheet sheet, String areaReference){
		return new NRange(sheet,Ranges.range(sheet.getNative(),areaReference));
	}
	
	public static NRange range(NSheet sheet, int tRow, int lCol, int bRow, int rCol){
		return new NRange(sheet,Ranges.range(sheet.getNative(),tRow,lCol,bRow,rCol));
	}
	
	public static NRange range(NSheet sheet, int row, int col){	
		return new NRange(sheet,Ranges.range(sheet.getNative(),row,col));
	}
}
