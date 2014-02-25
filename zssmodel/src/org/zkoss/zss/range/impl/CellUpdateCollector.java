package org.zkoss.zss.range.impl;

import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;

import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SheetRegion;

public class CellUpdateCollector {
	static ThreadLocal<CellUpdateCollector>  current = new ThreadLocal<CellUpdateCollector>();
	
	//a last object to prevent unnecessary cell-region creation
	SheetRegion last = null;
	
	private Set<SheetRegion> cellUpdates;
	
	public CellUpdateCollector(){
	}
	public static CellUpdateCollector setCurrent(CellUpdateCollector ctx){
		CellUpdateCollector old = current.get();
		current.set(ctx);
		return old;
	}
	
	public static CellUpdateCollector getCurrent(){
		return current.get();
	}

	public void addCellUpdate(SSheet sheet,int row, int column) {
		addCellUpdate(sheet,row,column,row,column);
	}
	public void addCellUpdate(SSheet sheet,int row, int column, int lastRow, int lastColumn) {
		if(last!=null && last.getSheet() == sheet && last.getRow() == row && last.getColumn()==column
				&& last.getLastRow() == lastRow && last.getLastColumn() == lastColumn){
			return;
		}
		addCellUpdate(new SheetRegion(sheet, row,column,lastRow,lastColumn));
	}
	public void addCellUpdate(SheetRegion cellUpdate) {
		if(cellUpdates==null){
			cellUpdates = new LinkedHashSet<SheetRegion>();
		}
		cellUpdates.add(cellUpdate);
		last = cellUpdate;
	}
	
	public Set<SheetRegion> getCellUpdates(){
		return cellUpdates==null?Collections.EMPTY_SET:Collections.unmodifiableSet(cellUpdates);
	}

	public void addCellUpdate(Set<SheetRegion> cellUpdates) {
		if(this.cellUpdates==null){
			this.cellUpdates = new LinkedHashSet<SheetRegion>();
		}
		this.cellUpdates.addAll(cellUpdates);
	}
}
