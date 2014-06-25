/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.range.impl;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.InvalidModelOpException;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SRow;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;

/**
 * 
 * @author dennis
 *
 */
//these code if from XRangeImpl,XBookHelper and migrate to new model
/*package*/ class AutoFilterHelper extends RangeHelperBase{

	public AutoFilterHelper(SRange range){
		super(range);
	}
	
	public CellRegion findAutoFilterRegion() {
		return new DataRegionHelper(range).findAutoFilterDataRegion();
	}
	
	
	//refer to #XRangeImpl#autoFilter
	public SAutoFilter enableAutoFilter(final boolean enable){
		SAutoFilter filter = sheet.getAutoFilter();
		boolean update = false;
		if(filter!=null && !enable){
			CellRegion region = filter.getRegion();
			SRange toUnhide = SRanges.range(sheet,region.getRow(),region.getColumn(),region.getLastRow(),region.getLastColumn()).getRows();
			//to show all hidden row in autofiler region when disable
			toUnhide.setHidden(false);
			sheet.deleteAutoFilter();
			filter = null;
			update = true;
		}else if(filter==null && enable){
			CellRegion region = findAutoFilterRegion();
			if(region!=null){
				filter = sheet.createAutoFilter(region);
				update = true;
			}else{
				throw new InvalidModelOpException("can't find any data in range");
			}
		}
		return filter;
	}
	
	//refer to #XRangeImpl#autoFilter(int field, Object criteria1, int filterOp, Object criteria2, Boolean visibleDropDown) {
	public SAutoFilter enableAutoFilter(final int field, final FilterOp filterOp,
			final Object criteria1, final Object criteria2, final Boolean showButton) {
		SAutoFilter filter = sheet.getAutoFilter();
		if(filter==null){
			CellRegion region = new DataRegionHelper(range).findAutoFilterDataRegion();
			if(region!=null){
				filter = sheet.createAutoFilter(region);
			}else{
				throw new InvalidModelOpException("can't find any data in range");
			}
		}
		
		final NFilterColumn fc = filter.getFilterColumn(field-1,true);	
		fc.setProperties(filterOp, criteria1, criteria2, showButton);

		//update rows
		final CellRegion affectedArea = filter.getRegion();
		final int row1 = affectedArea.getRow();
		final int col1 = affectedArea.getColumn(); 
		final int col =  col1 + field - 1;
		final int row = row1 + 1;
		final int row2 = affectedArea.getLastRow();
		final Set cr1 = fc.getCriteria1();
//		final Set<Ref> all = new HashSet<Ref>();
		for (int r = row; r <= row2; ++r) {
			final SCell cell = sheet.getCell(r, col); 
			final String val = isBlank(cell) ? "=" : getFormattedText(cell); //"=" means blank!
			if (cr1 != null && !cr1.isEmpty() && !cr1.contains(val)) { //to be hidden
				final SRow rowobj = sheet.getRow(r);
				if (!rowobj.isHidden()) { //a non-hidden row
					SRanges.range(sheet,r,0).getRows().setHidden(true);
				}
			} else { //candidate to be shown (other FieldColumn might still hide this row!
				final SRow rowobj = sheet.getRow(r);
				if (rowobj.isHidden() && canUnhide(filter, fc, r, col1)) { //a hidden row and no other hidden filtering
					// ZSS-646: we don't care about the columns at all; use 0.
//					final int left = sheet.getStartCellIndex(r);
//					final int right = sheet.getEndCellIndex(r);
					final SRange rng = SRanges.range(sheet,r,0,r,0);  
//					all.addAll(rng.getRefs());
					rng.getRows().setHidden(false); //unhide row
					
//					rng.notifyChange(); //why? text overflow? ->  //BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
				}
			}
		}
		
//		BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
				
		return filter;		
	}
	
	
	private boolean canUnhide(SAutoFilter af, NFilterColumn fc, int row, int col) {
		final Collection<NFilterColumn> fltcs = af.getFilterColumns();
		for(NFilterColumn fltc: fltcs) {
			if (fc.equals(fltc)) continue;
			if (shallHide(fltc, row, col)) { //any FilterColumn that shall hide the row
				return false;
			}
		}
		return true;
	}
	
	private boolean shallHide(NFilterColumn fc, int row, int col) {
		final SCell cell = sheet.getCell(row, col + fc.getIndex());
		final boolean blank = isBlank(cell); 
		final String val =  blank ? "=" : getFormattedText(cell); //"=" means blank!
		final Set critera1 = fc.getCriteria1();
		return critera1 != null && !critera1.isEmpty() && !critera1.contains(val);
	}

	//refer to XRangeImpl#showAllData
	public void resetAutoFilter() {
		SAutoFilter af = sheet.getAutoFilter();
		if (af == null) { //no AutoFilter to apply 
			return;
		}
		final CellRegion afrng = af.getRegion();
		final Collection<NFilterColumn> fcs = af.getFilterColumns();
		if (fcs == null)
			return;
		for(NFilterColumn fc : fcs) {
			fc.setProperties(FilterOp.VALUES, null, null, null); //clear all filter
		}
		final int row1 = afrng.getRow();
		final int row = row1 + 1;
		final int row2 = afrng.getLastRow();
		final int col1 = afrng.getColumn();
		final int col2 = afrng.getLastColumn();
//		final Set<Ref> all = new HashSet<Ref>();
		for (int r = row; r <= row2; ++r) {
			final SRow rowobj = sheet.getRow(r);
			if (rowobj.isHidden()) { //a hidden row
				//ZSS-646, we don't care about columns, use 0.
				//final int left = sheet.getStartCellIndex(r);
				//final int right = sheet.getEndCellIndex(r);
				final SRange rng = SRanges.range(sheet,r,0,r,0); 
//				all.addAll(rng.getRefs());
				rng.getRows().setHidden(false); //unhide
			}
		}

//		BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
		//update button
//		final XRangeImpl buttonChange = (XRangeImpl) XRanges.range(_sheet, row1, col1, row1, col2);
//		BookHelper.notifyBtnChanges(new HashSet<Ref>(buttonChange.getRefs()));
	}
	
	//refer to XRangeImpl#applyFilter
	public void applyAutoFilter() {
		SAutoFilter oldFilter = sheet.getAutoFilter();
		
		if (oldFilter==null) { //no criteria is applied
			return;
		}
		CellRegion region = oldFilter.getRegion();
		//copy filtering criteria
		int firstRow = region.getRow(); //first row is header
		int firstColumn = region.getColumn();
		//backup column data because getting from removed auto filter will cause XmlValueDisconnectedException
		
		//index,criteria1,op,criteria2,showVisible
		List<Object[]> originalFilteringColumns = new ArrayList<Object[]>();
		if (oldFilter.getFilterColumns() != null){ //has applied some criteria
			for (NFilterColumn filterColumn : oldFilter.getFilterColumns()){
				Object[] filterColumnData = new Object[5];
				filterColumnData[0] = filterColumn.getIndex()+1;
				filterColumnData[1] = filterColumn.getCriteria1().toArray(new String[0]);
				filterColumnData[2] = filterColumn.getOperator();
				filterColumnData[3] = filterColumn.getCriteria2();
				filterColumnData[4] = filterColumn.isShowButton();
				originalFilteringColumns.add(filterColumnData);
			}
		}
		
		enableAutoFilter(false); //disable existing filter
		
		//re-define filtering range 
		CellRegion filteringRange = DataRegionHelper.findCurrentRegion(sheet, firstRow, firstColumn);
		if (filteringRange == null){ //Don't enable auto filter if there are all blank cells
			return;
		}else{
			//enable auto filter
			sheet.createAutoFilter(filteringRange);
//			BookHelper.notifyAutoFilterChange(getRefs().iterator().next(),true);

			//apply original criteria
			for (int nCol = 0 ; nCol < originalFilteringColumns.size(); nCol ++){
				Object[] oldFilterColumn =  originalFilteringColumns.get(nCol);
				
				int field = (Integer)oldFilterColumn[0];
				Object c1 = oldFilterColumn[1];
				FilterOp op = (FilterOp)oldFilterColumn[2];
				Object c2 = oldFilterColumn[3];
				boolean showBtn = (Boolean)oldFilterColumn[4];
				
				enableAutoFilter(field,op,c1,c2,showBtn);
			}
		}	
			
	}

}
