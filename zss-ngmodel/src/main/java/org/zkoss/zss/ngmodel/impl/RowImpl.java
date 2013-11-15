package org.zkoss.zss.ngmodel.impl;

import java.util.HashMap;
import java.util.Iterator;
import java.util.TreeMap;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NRow;

public class RowImpl implements NRow {

	SheetImpl sheet;
	
	TreeMap<Integer,CellImpl> cells = new TreeMap<Integer,CellImpl>();
	HashMap<CellImpl,Integer> cellsReverse = new HashMap<CellImpl,Integer>(); 
	
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
			cellsReverse.put(cellObj, columnIdx);
			
			//create column since we have cell of it
			sheet.getOrCreateColumnAt(columnIdx);
		}
		return cellObj;
	}
	
	int getCellIndex(CellImpl cell){
		Integer index = cellsReverse.get(cell);
		return index==null?-1:index.intValue();
	}

	public int getStartCellIndex() {
		Integer k = cells.isEmpty()?null:cells.firstKey();
		return k==null?-1:k.intValue();
	}

	public int getEndCellIndex() {
		Integer k = cells.isEmpty()?null:cells.lastKey();
		return k==null?-1:k.intValue();
	}

	protected void onModelEvent(ModelEvent event) {
		// TODO Auto-generated method stub
		
	}

	public void clearCell(int start, int end) {
		start = Math.max(start,getStartCellIndex());
		end = Math.min(end,getEndCellIndex());
		
		//loop from start to end, or from iteration? which one is better?
		if( end-start > cells.size() ){
			Iterator<Integer> iter = cells.keySet().iterator();
			while(iter.hasNext()){
				int idx = iter.next();
				if(idx>=start && idx<=end){
					cellsReverse.remove(cells.get(idx));
					iter.remove();
				}
			}
		}else{
			for(int i=start;i<=end;i++){
				CellImpl cell = cells.remove(i);
				if(cell!=null){
					cellsReverse.remove(cell);
				}
			}
		}
	}

}
