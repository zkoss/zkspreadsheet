package org.zkoss.zss.ngmodel.impl;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.NavigableMap;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;

import javax.swing.text.NavigationFilter;

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
	
	BiIndexPool<RowImpl> rows = new BiIndexPool<RowImpl>();
	BiIndexPool<ColumnImpl> columns = new BiIndexPool<ColumnImpl>();
	
	
	public SheetImpl(BookImpl book){
		this.book = book;
	}
	
	public NBook getBook() {
		return book;
	}

	public String getSheetName() {
		return name;
	}

	public NRow getRow(int rowIdx) {
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
		}
		return rowObj;
	}
	
	int getRowIndex(RowImpl row){
		return rows.get(row);
	}

	public NColumn getColumn(int columnIdx) {
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
		}
		return columnObj;
	}
	
	int getColumnIndex(ColumnImpl column){
		return columns.get(column);
	}

	public NCell getCell(int rowIdx, int columnIdx) {
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

	public int getStartRowIndex() {
		return rows.firstKey();
	}

	public int getEndRowIndex() {
		return rows.lastKey();
	}
	
	public int getStartColumnIndex() {
		return columns.firstKey();
	}

	public int getEndColumnIndex() {
		return columns.lastKey();
	}

	public int getStartColumnIndex(int row) {
		RowImpl rowObj = (RowImpl) getRowAt(row,false);
		if(rowObj!=null){
			return rowObj.getStartCellIndex();
		}
		return -1;
	}

	public int getEndColumn(int row) {
		RowImpl rowObj = (RowImpl) getRowAt(row,false);
		if(rowObj!=null){
			return rowObj.getEndCellIndex();
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
	

	public void clearRow(int rowIdx, int rowIdx2) {
		int start = Math.max(Math.min(rowIdx, rowIdx2),getStartRowIndex());
		int end = Math.min(Math.max(rowIdx, rowIdx2),getEndRowIndex());
				
		rows.clear(start,end);
		
		//Send event?
		
	}

	public void clearColumn(int columnIdx, int columnIdx2) {
		int start = Math.max(Math.min(columnIdx, columnIdx2),getStartColumnIndex());
		int end = Math.min(Math.max(columnIdx, columnIdx2),getEndColumnIndex());
		
		columns.clear(start,end);
		
		for(RowImpl row:rows.values()){
			row.clearCell(start,end);
		}
		//Send event?
		
	}

	public void clearCell(int rowIdx, int columnIdx, int rowIdx2,
			int columnIdx2) {
		int rowStart = Math.max(Math.min(rowIdx, rowIdx2),getStartRowIndex());
		int rowEnd = Math.min(Math.max(rowIdx, rowIdx2),getEndRowIndex());
		int columnStart = Math.max(Math.min(columnIdx, columnIdx2),getStartColumnIndex());
		int columnEnd = Math.min(Math.max(columnIdx, columnIdx2),getEndColumnIndex());
		
		Collection<RowImpl> effected = rows.subValues(rowStart,rowEnd);
		
		Iterator<RowImpl> iter = effected.iterator();
		while(iter.hasNext()){
			iter.next().clearCell(columnStart, columnEnd);
		}
	}

	public void insertRow(int rowIdx, int size) {
		if(size<=0) return;
		
		int end = getEndRowIndex();
		if(rowIdx>end) return;
		
		int start = Math.max(rowIdx,getStartRowIndex());
		
		rows.insert(start, size);
		
		//Event
		
	}

	public void deleteRow(int rowIdx, int size) {
		if(size<=0) return;
		
		int end = getEndRowIndex();
		if(rowIdx>end) return;
		
		int start = Math.max(rowIdx,getStartRowIndex());
		
		rows.delete(start, size);
		
		//Event
	}
	

	protected void copySheet(SheetImpl sheet) {
		//can only clone on the begining.
		
		//TODO
		throw new UnsupportedOperationException("not implement yet");
	}

	public void dump(StringBuilder builder) {
		
		builder.append("'").append(getSheetName()).append("' {\n");
		
		int endColumn = getEndColumnIndex();
		int endRow = getEndRowIndex();
		builder.append("  ==Columns==\n\t");
		for(int i=0;i<=endColumn;i++){
			builder.append(i).append("\t");
		}
		builder.append("\n");
		builder.append("  ==Row==");
		for(int i=0;i<=endRow;i++){
			builder.append("\n  ").append(i).append("\t");
			if(getRow(i).isNull()){
				builder.append("-*");
				continue;
			}
			for(int j=0;j<=endColumn;j++){
				NCell cell = getCell(i, j);
				Object cellvalue = cell.isNull()?"-":cell.getValue();
				String str = cellvalue==null?"null":cellvalue.toString();
				if(str.length()>8){
					str = str.substring(0,8);
				}else{
					str = str+"\t";
				}
				
				builder.append(str);
			}
		}
		builder.append("}\n");
	}

	public void insertColumn(int columnIdx, int size) {
		if(size<=0) return;
		
		int end = getEndColumnIndex();
		if(columnIdx>end) return;
		
		int start = Math.max(columnIdx,getStartColumnIndex());
		
		columns.insert(start, size);
		
		for(RowImpl row:rows.values()){
			row.insertCell(start,size);
		}
		//Send event?
		
	}

	public void deleteColumn(int columnIdx, int size) {
		if(size<=0) return;
		
		int end = getEndColumnIndex();
		if(columnIdx>end) return;
		
		int start = Math.max(columnIdx,getStartColumnIndex());
		
		columns.delete(start, size);
		
		for(RowImpl row:rows.values()){
			row.deleteCell(start,size);
		}
		//Send event?
	}



}
