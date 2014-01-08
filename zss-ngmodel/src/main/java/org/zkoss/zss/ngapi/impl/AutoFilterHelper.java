package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.NRanges;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.NAutoFilter;
import org.zkoss.zss.ngmodel.NSheet;


public class AutoFilterHelper {

	private NRange range;
	
	public AutoFilterHelper(NRange range){
		this.range = range;
	}

//	public CellRegion clearAutoFilter() {
//		NSheet sheet = range.getSheet();
//		NAutoFilter filter = sheet.getAutoFilter();
//		if(filter==null)
//			return null;
//		
//		CellRegion region = filter.getRegion();
//		if(region==null)
//			return null;
//		
//		NRange toUnhide = NRanges.range(sheet,range.getRow(),range.getColumn(),range.getLastRow(),range.getLastColumn()).getRows();
//		toUnhide.setHidden(false);
//		sheet.deleteAutoFilter();
//		
//		return region;
//	}
//
//	public CellRegion createAutoFilter() {
//		NSheet sheet = range.getSheet();
//		NAutoFilter filter = sheet.getAutoFilter();
//		if(filter!=null)
//			return filter.getRegion();
//		
//		NRange filterRange = this.range.findAutoFilterRange();
//		if(filterRange==null){
//			return null;
//		}
//		CellRegion region = new CellRegion(filterRange.getRow(),filterRange.getColumn(),filterRange.getLastColumn(),filterRange.getLastColumn());
//		filter = sheet.createAutoFilter(region);
//		return region;
//	}
}
