package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NRow;

public class RowImpl implements NRow {

	SheetImpl sheet;
	
	BiIndexPool<CellImpl> cells = new BiIndexPool<CellImpl>(); 
	
	public RowImpl(SheetImpl sheet) {
		this.sheet = sheet;
	}

	public int getIndex() {
		return sheet.getRowIndex(this);
	}

	public boolean isNull() {
		return false;
	}
	
	public NCell getCellAt(int columnIdx) {
		return getCellAt(columnIdx,true);
	}
	
	NCell getCellAt(int columnIdx, boolean proxy) {
		NCell cellObj = cells.get(columnIdx);
		if(cellObj != null){
			return cellObj;
		}
		return proxy?new CellProxyImpl(sheet,getIndex(),columnIdx):null;
	}
	
	NCell getOrCreateCellAt(int columnIdx){
		CellImpl cellObj = cells.get(columnIdx);
		if(cellObj == null){
			cellObj = new CellImpl(this);
			cells.put(columnIdx, cellObj);
			//create column since we have cell of it
			sheet.getOrCreateColumnAt(columnIdx);
		}
		return cellObj;
	}
	
	int getCellIndex(CellImpl cell){
		return cells.get(cell);
	}

	public int getStartCellIndex() {
		return cells.firstKey();
	}

	public int getEndCellIndex() {
		return cells.lastKey();
	}

	protected void onModelEvent(ModelEvent event) {
		// TODO Auto-generated method stub
		
	}

	public void clearCell(int start, int end) {
		start = Math.max(start,getStartCellIndex());
		end = Math.min(end,getEndCellIndex());
		
		cells.clear(start, end);
	}

	public void insertCell(int cellIdx, int size) {
		if(size<=0) return;
		
		int end = getEndCellIndex();
		if(cellIdx>end) return;
		
		int start = Math.max(cellIdx,getStartCellIndex());
		
		cells.insert(start, size);
	}

	public void deleteCell(int cellIdx, int size) {
		if(size<=0) return;
		
		int end = getEndCellIndex();
		if(cellIdx>end) return;
		
		int start = Math.max(cellIdx,getStartCellIndex());
		
		cells.delete(start, size);
	}

}
