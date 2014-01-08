package org.zkoss.zss.ngapi.impl;

import java.util.Collection;
import java.util.HashSet;
import java.util.Set;

import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.NRanges;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.NAutoFilter;
import org.zkoss.zss.ngmodel.NAutoFilter.FilterOp;
import org.zkoss.zss.ngmodel.NAutoFilter.NFilterColumn;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.format.FormatEngine;

/**
 * 
 * @author dennis
 *
 */
//these code if from XRangeImpl,XBookHelper and migrate to new model
/*package*/ class AutoFilterHelper extends RangeHelperBase{

	public AutoFilterHelper(NRange range){
		super(range);
	}
	
	public CellRegion findAutoFilterRegion() {
		return new DataRegionHelper(range).findAutoFilterDataRegion();
	}
	
	
	//refer to #XRangeImpl#autoFilter
	public NAutoFilter enableAutoFilter(final boolean enable){
		NAutoFilter filter = sheet.getAutoFilter();
		boolean update = false;
		if(filter!=null && !enable){
			CellRegion region = filter.getRegion();
			NRange toUnhide = NRanges.range(sheet,region.getRow(),region.getColumn(),region.getLastRow(),region.getLastColumn()).getRows();
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
				throw new IllegalStateException("null region, you should call #findAutoFilterRange to ensure a valid region before enalbe autofilter");
			}
		}
		return filter;
	}
	
	//refer to #XRangeImpl#autoFilter(int field, Object criteria1, int filterOp, Object criteria2, Boolean visibleDropDown) {
	public NAutoFilter enableAutoFilter(final int field, final FilterOp filterOp,
			final Object criteria1, final Object criteria2, final Boolean visibleDropDown) {
		NAutoFilter filter = sheet.getAutoFilter();
		if(filter==null){
			CellRegion region = new DataRegionHelper(range).findAutoFilterDataRegion();
			if(region!=null){
				filter = sheet.createAutoFilter(region);
			}else{
				throw new IllegalStateException("null region, you should call #findAutoFilterRange to ensure a valid region before enalbe autofilter");
			}
		}
		
		final NFilterColumn fc = filter.getFilterColumn(field-1,true);	
		fc.setProperties(filterOp, criteria1, criteria2, visibleDropDown);

		//update rows
		final CellRegion affectedArea = filter.getRegion();
		final int row1 = affectedArea.getRow();
		final int col1 = affectedArea.getColumn(); 
		final int col =  col1 + field - 1;
		final int row = row1 + 1;
		final int row2 = affectedArea.getLastRow();
		final Set cr1 = fc.getCriteria1();
		final Set<Ref> all = new HashSet<Ref>();
		for (int r = row; r <= row2; ++r) {
			final NCell cell = sheet.getCell(r, col); 
			final String val = isBlank(cell) ? "=" : getFormattedText(cell); //"=" means blank!
			if (cr1 != null && !cr1.isEmpty() && !cr1.contains(val)) { //to be hidden
				final NRow rowobj = sheet.getRow(r);
				if (!rowobj.isHidden()) { //a non-hidden row
					NRanges.range(sheet,r,0).getRows().setHidden(true);
				}
			} else { //candidate to be shown (other FieldColumn might still hide this row!
				final NRow rowobj = sheet.getRow(r);
				if (rowobj.isHidden() && canUnhide(filter, fc, r, col1)) { //a hidden row and no other hidden filtering
					final int left = sheet.getStartCellIndex(r);
					final int right = sheet.getEndCellIndex(r);
					final NRange rng = NRanges.range(sheet,r,left,r,right); 
//					all.addAll(rng.getRefs());
					rng.getRows().setHidden(false); //unhide row
					
//					rng.notifyChange(); //why? text overflow? ->  //BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
				}
			}
		}
		
//		BookHelper.notifyCellChanges(_sheet.getBook(), all); //unhidden row must reevaluate
				
		return filter;		
	}
	
	
	private boolean canUnhide(NAutoFilter af, NFilterColumn fc, int row, int col) {
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
		final NCell cell = sheet.getCell(row, col + fc.getIndex());
		final boolean blank = isBlank(cell); 
		FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
		final String val =  blank ? "=" : getFormattedText(cell); //"=" means blank!
		final Set critera1 = fc.getCriteria1();
		return critera1 != null && !critera1.isEmpty() && !critera1.contains(val);
	}

}
