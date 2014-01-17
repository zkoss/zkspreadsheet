package org.zkoss.zss.ngapi.impl;

import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;

import org.zkoss.zss.ngmodel.CellRegion;

public class CellUpdateCollector {
	static ThreadLocal<CellUpdateCollector>  current = new ThreadLocal<CellUpdateCollector>();
	
	//a last object to prevent unnecessary cell-region creation
	CellRegion last = null;
	
	private Set<CellRegion> cellUpdates;
	
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

	public void addCellUpdate(int row, int column) {
		addCellUpdate(row,column,row,column);
	}
	public void addCellUpdate(int row, int column, int lastRow, int lastColumn) {
		if(last!=null && last.row == row && last.column==column
				&& last.lastRow == lastRow && last.lastColumn == lastColumn){
			return;
		}
		addCellUpdate(new CellRegion(row,column,lastRow,lastColumn));
	}
	public void addCellUpdate(CellRegion cellUpdate) {
		if(cellUpdates==null){
			cellUpdates = new LinkedHashSet<CellRegion>();
		}
		cellUpdates.add(cellUpdate);
		last = cellUpdate;
	}
	
	public Set<CellRegion> getCellUpdates(){
		return cellUpdates==null?Collections.EMPTY_SET:Collections.unmodifiableSet(cellUpdates);
	}

	public void addCellUpdate(Set<CellRegion> cellUpdates) {
		if(this.cellUpdates==null){
			this.cellUpdates = new LinkedHashSet<CellRegion>();
		}
		this.cellUpdates.addAll(cellUpdates);
	}
}
