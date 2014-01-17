package org.zkoss.zss.ngapi.impl;

import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;

import org.zkoss.zss.ngmodel.CellRegion;

public class CellUpdateCollector {
	static ThreadLocal<CellUpdateCollector>  current = new ThreadLocal<CellUpdateCollector>();
	
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
		addCellUpdate(new CellRegion(row,column));
	}
	public void addCellUpdate(int row, int column, int lastRow, int lastColumn) {
		addCellUpdate(new CellRegion(row,column,lastRow,lastColumn));
	}
	public void addCellUpdate(CellRegion cellUpdate) {
		if(cellUpdates==null){
			cellUpdates = new LinkedHashSet<CellRegion>();
		}
		cellUpdates.add(cellUpdate);
	}
	
	public Set<CellRegion> getCellUpdate(){
		return cellUpdates==null?Collections.EMPTY_SET:Collections.unmodifiableSet(cellUpdates);
	}

	public void addCellUpdate(Set<CellRegion> cellUpdates) {
		if(this.cellUpdates==null){
			this.cellUpdates = new LinkedHashSet<CellRegion>();
		}
		this.cellUpdates.addAll(cellUpdates);
	}
}
