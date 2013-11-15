package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.util.CellReference;

public class RowImpl implements NRow {

	SheetImpl sheet;
	
	BiIndexPool<CellImpl> cells = new BiIndexPool<CellImpl>();
	
	CellStyleImpl cellStyle;
	
	public RowImpl(SheetImpl sheet) {
		this.sheet = sheet;
	}

	public int getIndex() {
		checkOrphan();
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
		checkOrphan();
		return proxy?new CellProxyImpl(sheet,getIndex(),columnIdx):null;
	}
	
	NCell getOrCreateCellAt(int columnIdx){
		CellImpl cellObj = cells.get(columnIdx);
		if(cellObj == null){
			checkOrphan();
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
		
		for(CellImpl cell:cells.clear(start, end)){
			cell.release();
		}
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
		
		for(CellImpl cell:cells.delete(start, size)){
			cell.release();
		}
	}
	
	public String asString() {
		return String.valueOf(getIndex()+1);
	}
	
	protected void checkOrphan(){
		if(sheet==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
	protected void release(){
		sheet = null;
	}

	public NCellStyle getCellStyle() {
		return cellStyle;
	}

}
