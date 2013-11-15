package org.zkoss.zss.ngmodel.impl;

import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NViewInfo;

public class SheetImpl implements NSheet {

	BookImpl book;
	String name;
	
	TreeMap<Integer,RowImpl> rows = new TreeMap<Integer,RowImpl>();
	HashMap<RowImpl,Integer> rowsReverse = new HashMap<RowImpl,Integer>(); 
	
	TreeMap<Integer,ColumnImpl> columns = new TreeMap<Integer,ColumnImpl>();
	HashMap<ColumnImpl,Integer> columnsReverse = new HashMap<ColumnImpl,Integer>();
	
	
	public SheetImpl(BookImpl book){
		this.book = book;
	}
	
	public NBook getBook() {
		return book;
	}

	public String getSheetName() {
		return name;
	}

	public NRow getRowAt(int rowIdx) {
		return getRowAt(rowIdx,true);
	}
	
	NRow getRowAt(int rowIdx, boolean proxy) {
		NRow rowObj = rows.get(rowIdx);
		if(rowObj != null){
			return rowObj;
		}
		return proxy?new RowProxyImpl(this,rowIdx):null;
	}
	
	NRow getOrCreateRowAt(int rowIdx){
		RowImpl rowObj = rows.get(rowIdx);
		if(rowObj == null){
			rowObj = new RowImpl(this);
			rows.put(rowIdx, rowObj);
			rowsReverse.put(rowObj, rowIdx);
		}
		return rowObj;
	}
	
	int getRowIndex(RowImpl row){
		Integer index = rowsReverse.get(row);
		return index==null?-1:index.intValue();
	}

	public NColumn getColumnAt(int columnIdx) {
		return getColumnAt(columnIdx,true);
	}
	
	NColumn getColumnAt(int columnIdx, boolean proxy) {
		NColumn colObj = columns.get(columnIdx);
		if(colObj != null){
			return colObj;
		}
		return proxy?new ColumnProxyImpl(this,columnIdx):null;
	}
	
	NColumn getOrCreateColumnAt(int columnIdx){
		ColumnImpl columnObj = columns.get(columnIdx);
		if(columnObj == null){
			columnObj = new ColumnImpl(this);
			columns.put(columnIdx, columnObj);
			columnsReverse.put(columnObj, columnIdx);
		}
		return columnObj;
	}
	
	int getColumnIndex(ColumnImpl column){
		Integer index = columnsReverse.get(column);
		return index==null?-1:index.intValue();
	}

	public NCell getCellAt(int rowIdx, int columnIdx) {
		return getCellAt(rowIdx,columnIdx,true);
	}
	
	public NCell getCellAt(int rowIdx, int columnIdx, boolean proxy) {
		RowImpl rowObj = (RowImpl) getRowAt(rowIdx,false);
		if(rowObj!=null){
			return rowObj.getCellAt(columnIdx,proxy);
		}
		return proxy?new CellProxyImpl(this, rowIdx,columnIdx):null;
	}
	
	NCell getOrCreateCellAt(int rowIdx, int columnIdx){
		RowImpl rowObj = (RowImpl)getOrCreateRowAt(rowIdx);
		NCell cell = rowObj.getOrCreateCellAt(columnIdx);
		return cell;
	}

	public int getStartRow() {
		Integer k = rows.isEmpty()?null:rows.firstKey();
		return k==null?-1:k.intValue();
	}

	public int getEndRow() {
		Integer k = rows.isEmpty()?null:rows.lastKey();
		return k==null?-1:k.intValue();
	}
	
	public int getStartColumn() {
		Integer k = columns.isEmpty()?null:columns.firstKey();
		return k==null?-1:k.intValue();
	}

	public int getEndColumn() {
		Integer k = columns.isEmpty()?null:columns.lastKey();
		return k==null?-1:k.intValue();
	}

	public int getStartColumn(int row) {
		RowImpl rowObj = (RowImpl) getRowAt(row,false);
		if(rowObj!=null){
			return rowObj.getStartColumn();
		}
		return -1;
	}

	public int getEndColumn(int row) {
		RowImpl rowObj = (RowImpl) getRowAt(row,false);
		if(rowObj!=null){
			return rowObj.getEndColumn();
		}
		return -1;
	}


	public void setSheetName(String name) {
		this.name = name;
	}

	protected void onModelEvent(ModelEvent event) {
		for(RowImpl row:rows.values()){
			row.onModelEvent(event);
		}
		for(ColumnImpl column:columns.values()){
			column.onModelEvent(event);
		}
		//TODO to other object
	}

	protected void copySheet(SheetImpl sheet) {
		//can only clone on the begining.
		
		//TODO
		throw new UnsupportedOperationException("not implement yet");
	}

	public void dump(StringBuilder builder) {
		int endColumn = getEndColumn();
		int endRow = getEndRow();
		builder.append("==Columns==\n\t");
		for(int i=0;i<=endColumn;i++){
			builder.append(i).append("\t");
		}
		builder.append("\n");
		builder.append("==Row==\n");
		for(int i=0;i<=endRow;i++){
			builder.append(i).append("\t");
			for(int j=0;j<=endColumn;j++){
				NCell cell = getCellAt(i, j);
				builder.append(cell.getValue()).append("\t");
			}
			builder.append("\n");
		}
	}

}
