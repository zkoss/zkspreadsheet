package org.zkoss.zss.ngmodel.impl;

import java.util.Collection;
import java.util.Iterator;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.util.CellReference;

public class SheetImpl extends AbstractSheet {
	private static final long serialVersionUID = 1L;
	private AbstractBook book;
	private String name;
	private final String id;
	
	private final BiIndexPool<AbstractRow> rows = new BiIndexPool<AbstractRow>();
	private final BiIndexPool<AbstractColumn> columns = new BiIndexPool<AbstractColumn>();
	
	
	public SheetImpl(AbstractBook book,String id){
		this.book = book;
		this.id = id;
	}
	
	public NBook getBook() {
		checkOrphan();
		return book;
	}

	public String getSheetName() {
		return name;
	}

	public NRow getRow(int rowIdx) {
		return getRowAt(rowIdx,true);
	}
	@Override
	AbstractRow getRowAt(int rowIdx, boolean proxy) {
		AbstractRow rowObj = rows.get(rowIdx);
		if(rowObj != null){
			return rowObj;
		}
		return proxy?new RowProxyImpl(this,rowIdx):null;
	}
	@Override
	AbstractRow getOrCreateRowAt(int rowIdx){
		AbstractRow rowObj = rows.get(rowIdx);
		if(rowObj == null){
			rowObj = new RowImpl(this);
			rows.put(rowIdx, rowObj);
		}
		return rowObj;
	}
	@Override
	int getRowIndex(AbstractRow row){
		return rows.get(row);
	}

	public NColumn getColumn(int columnIdx) {
		return getColumnAt(columnIdx,true);
	}
	@Override
	AbstractColumn getColumnAt(int columnIdx, boolean proxy) {
		AbstractColumn colObj = columns.get(columnIdx);
		if(colObj != null){
			return colObj;
		}
		return proxy?new ColumnProxyImpl(this,columnIdx):null;
	}
	@Override
	AbstractColumn getOrCreateColumnAt(int columnIdx){
		AbstractColumn columnObj = columns.get(columnIdx);
		if(columnObj == null){
			columnObj = new ColumnImpl(this);
			columns.put(columnIdx, columnObj);
		}
		return columnObj;
	}
	@Override
	int getColumnIndex(AbstractColumn column){
		return columns.get(column);
	}

	public NCell getCell(int rowIdx, int columnIdx) {
		return getCellAt(rowIdx,columnIdx,true);
	}
	
	@Override
	AbstractCell getCellAt(int rowIdx, int columnIdx, boolean proxy) {
		AbstractRow rowObj = (AbstractRow) getRowAt(rowIdx,false);
		if(rowObj!=null){
			return rowObj.getCellAt(columnIdx,proxy);
		}
		return proxy?new CellProxyImpl(this, rowIdx,columnIdx):null;
	}
	@Override
	AbstractCell getOrCreateCellAt(int rowIdx, int columnIdx){
		AbstractRow rowObj = (AbstractRow)getOrCreateRowAt(rowIdx);
		AbstractCell cell = rowObj.getOrCreateCellAt(columnIdx);
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
		AbstractRow rowObj = (AbstractRow) getRowAt(row,false);
		if(rowObj!=null){
			return rowObj.getStartCellIndex();
		}
		return -1;
	}

	public int getEndColumn(int row) {
		AbstractRow rowObj = (AbstractRow) getRowAt(row,false);
		if(rowObj!=null){
			return rowObj.getEndCellIndex();
		}
		return -1;
	}

	@Override
	void setSheetName(String name) {
		this.name = name;
	}
	@Override
	void onModelEvent(ModelEvent event) {
		for(AbstractRow row:rows.values()){
			row.onModelEvent(event);
		}
		for(AbstractColumn column:columns.values()){
			column.onModelEvent(event);
		}
		//TODO to other object
	}
	
	public void clearRow(int rowIdx, int rowIdx2) {
		int start = Math.max(Math.min(rowIdx, rowIdx2),getStartRowIndex());
		int end = Math.min(Math.max(rowIdx, rowIdx2),getEndRowIndex());
				
		for(AbstractRow row:rows.clear(start,end)){
			row.release();
		}
		
		//Send event?
		
	}

	public void clearColumn(int columnIdx, int columnIdx2) {
		int start = Math.max(Math.min(columnIdx, columnIdx2),getStartColumnIndex());
		int end = Math.min(Math.max(columnIdx, columnIdx2),getEndColumnIndex());
		
		for(AbstractColumn column:columns.clear(start,end)){
			column.release();
		}
		
		for(AbstractRow row:rows.values()){
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
		
		Collection<AbstractRow> effected = rows.subValues(rowStart,rowEnd);
		
		Iterator<AbstractRow> iter = effected.iterator();
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
		
		for(AbstractRow row:rows.delete(start, size)){
			row.release();
		}
		
		//Event
	}
	
	@Override
	void copyTo(AbstractSheet sheet) {
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
			builder.append(CellReference.convertNumToColString(i)).append(":").append(i).append("\t");
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
		
		for(AbstractRow row:rows.values()){
			row.insertCell(start,size);
		}
		//Send event?
		
	}

	public void deleteColumn(int columnIdx, int size) {
		if(size<=0) return;
		
		int end = getEndColumnIndex();
		if(columnIdx>end) return;
		
		int start = Math.max(columnIdx,getStartColumnIndex());
		
		for(AbstractColumn column:columns.delete(start, size)){
			column.release();
		}
		
		for(AbstractRow row:rows.values()){
			row.deleteCell(start,size);
		}
		//Send event?
	}

	
	protected void checkOrphan(){
		if(book==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
	@Override
	void release(){
		book = null;
	}

	public String getId() {
		return id;
	}
	
	

}
