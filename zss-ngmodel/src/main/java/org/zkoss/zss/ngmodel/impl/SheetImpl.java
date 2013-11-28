package org.zkoss.zss.ngmodel.impl;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import org.zkoss.zss.ngmodel.InvalidateModelOpException;
import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NChart;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NPicture;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NChart.NChartType;
import org.zkoss.zss.ngmodel.NPicture.Format;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.NViewAnchor;
import org.zkoss.zss.ngmodel.chart.NChartData;
import org.zkoss.zss.ngmodel.impl.chart.CategoryChartDataImpl;
import org.zkoss.zss.ngmodel.util.CellReference;

public class SheetImpl extends SheetAdv {
	private static final long serialVersionUID = 1L;
	private BookAdv book;
	private String name;
	private final String id;
	
	private final BiIndexPool<RowAdv> rows = new BiIndexPool<RowAdv>();
	private final BiIndexPool<ColumnAdv> columns = new BiIndexPool<ColumnAdv>();
	
	private final List<PictureAdv> pictures = new LinkedList<PictureAdv>();
	private final List<ChartAdv> charts = new LinkedList<ChartAdv>();
	
	
	public SheetImpl(BookAdv book,String id){
		this.book = book;
		this.id = id;
	}
	
	protected void checkOwnership(NPicture picture){
		if(!pictures.contains(picture)){
			throw new InvalidateModelOpException("doesn't has ownership "+ picture);
		}
	}
	
	protected void checkOwnership(NChart chart){
		if(!charts.contains(chart)){
			throw new InvalidateModelOpException("doesn't has ownership "+ chart);
		}
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
	RowAdv getRowAt(int rowIdx, boolean proxy) {
		RowAdv rowObj = rows.get(rowIdx);
		if(rowObj != null){
			return rowObj;
		}
		return proxy?new RowProxy(this,rowIdx):null;
	}
	@Override
	RowAdv getOrCreateRowAt(int rowIdx){
		RowAdv rowObj = rows.get(rowIdx);
		if(rowObj == null){
			rowObj = new RowImpl(this);
			rows.put(rowIdx, rowObj);
		}
		return rowObj;
	}
	@Override
	int getRowIndex(RowAdv row){
		return rows.get(row);
	}

	public NColumn getColumn(int columnIdx) {
		return getColumnAt(columnIdx,true);
	}
	@Override
	ColumnAdv getColumnAt(int columnIdx, boolean proxy) {
		ColumnAdv colObj = columns.get(columnIdx);
		if(colObj != null){
			return colObj;
		}
		return proxy?new ColumnProxy(this,columnIdx):null;
	}
	@Override
	ColumnAdv getOrCreateColumnAt(int columnIdx){
		ColumnAdv columnObj = columns.get(columnIdx);
		if(columnObj == null){
			columnObj = new ColumnImpl(this);
			columns.put(columnIdx, columnObj);
		}
		return columnObj;
	}
	@Override
	int getColumnIndex(ColumnAdv column){
		return columns.get(column);
	}

	public NCell getCell(int rowIdx, int columnIdx) {
		return getCellAt(rowIdx,columnIdx,true);
	}
	
	@Override
	CellAdv getCellAt(int rowIdx, int columnIdx, boolean proxy) {
		RowAdv rowObj = (RowAdv) getRowAt(rowIdx,false);
		if(rowObj!=null){
			return rowObj.getCellAt(columnIdx,proxy);
		}
		return proxy?new CellProxy(this, rowIdx,columnIdx):null;
	}
	@Override
	CellAdv getOrCreateCellAt(int rowIdx, int columnIdx){
		RowAdv rowObj = (RowAdv)getOrCreateRowAt(rowIdx);
		CellAdv cell = rowObj.getOrCreateCellAt(columnIdx);
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
		RowAdv rowObj = (RowAdv) getRowAt(row,false);
		if(rowObj!=null){
			return rowObj.getStartCellIndex();
		}
		return -1;
	}

	public int getEndColumn(int row) {
		RowAdv rowObj = (RowAdv) getRowAt(row,false);
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
		for(RowAdv row:rows.values()){
			row.onModelEvent(event);
		}
		for(ColumnAdv column:columns.values()){
			column.onModelEvent(event);
		}
		//TODO to other object
	}
	
	public void clearRow(int rowIdx, int rowIdx2) {
		int start = Math.min(rowIdx, rowIdx2);
		int end = Math.max(rowIdx, rowIdx2);
		
		//clear before move relation
		for(RowAdv row:rows.subValues(start,end)){
			row.release();
		}		
		rows.clear(start,end);
		
		//Send event?
		
	}

	public void clearColumn(int columnIdx, int columnIdx2) {
		int start = Math.min(columnIdx, columnIdx2);
		int end = Math.max(columnIdx, columnIdx2);
		
		
		for(ColumnAdv column:columns.subValues(start,end)){
			column.release();
		}
		columns.clear(start,end);
		
		for(RowAdv row:rows.values()){
			row.clearCell(start,end);
		}
		//Send event?
		
	}

	public void clearCell(int rowIdx, int columnIdx, int rowIdx2,
			int columnIdx2) {
		int rowStart = Math.min(rowIdx, rowIdx2);
		int rowEnd = Math.max(rowIdx, rowIdx2);
		int columnStart = Math.min(columnIdx, columnIdx2);
		int columnEnd = Math.max(columnIdx, columnIdx2);
		
		Collection<RowAdv> effected = rows.subValues(rowStart,rowEnd);
		
		Iterator<RowAdv> iter = effected.iterator();
		while(iter.hasNext()){
			iter.next().clearCell(columnStart, columnEnd);
		}
	}

	public void insertRow(int rowIdx, int size) {
		checkOrphan();
		if(size<=0) return;
		
		rows.insert(rowIdx, size);
		
		shiftAfterRowInsert(rowIdx,size);
		
		book.sendEvent(ModelEvents.ON_ROW_INSERTED, ModelEvents.PARAM_SHEET, this, ModelEvents.PARAM_ROW_INDEX, rowIdx, 
				ModelEvents.PARAM_SIZE, size);
	}
	
	private void shiftAfterRowInsert(int rowIdx, int size) {
		// handling pic shift
		for (PictureAdv pic : pictures) {
			NViewAnchor anchor = pic.getAnchor();
			int idx = anchor.getRowIndex();
			if (idx >= rowIdx) {
				anchor.setRowIndex(idx + size);
			}
		}
		// handling pic shift
		for (ChartAdv chart : charts) {
			NViewAnchor anchor = chart.getAnchor();
			int idx = anchor.getRowIndex();
			if (idx >= rowIdx) {
				anchor.setRowIndex(idx + size);
			}
		}
	}
	private void shiftAfterRowDelete(int rowIdx, int size) {
		//handling pic shift
		for(PictureAdv pic:pictures){
			NViewAnchor anchor = pic.getAnchor();
			int idx = anchor.getRowIndex();
			if(idx >= rowIdx+size){
				anchor.setRowIndex(idx-size);
			}else if(idx >= rowIdx){
				anchor.setRowIndex(rowIdx);//as excel's rule
				anchor.setYOffset(0);
			}
		}
		//handling pic shift
		for(ChartAdv chart:charts){
			NViewAnchor anchor = chart.getAnchor();
			int idx = anchor.getRowIndex();
			if(idx >= rowIdx+size){
				anchor.setRowIndex(idx-size);
			}else if(idx >= rowIdx){
				anchor.setRowIndex(rowIdx);//as excel's rule
				anchor.setYOffset(0);
			}
		}			
	}
	private void shiftAfterColumnInsert(int columnIdx, int size) {
		// handling pic shift
		for (PictureAdv pic : pictures) {
			NViewAnchor anchor = pic.getAnchor();
			int idx = anchor.getColumnIndex();
			if (idx >= columnIdx) {
				anchor.setColumnIndex(idx + size);
			}
		}
		// handling pic shift
		for (ChartAdv chart : charts) {
			NViewAnchor anchor = chart.getAnchor();
			int idx = anchor.getColumnIndex();
			if (idx >= columnIdx) {
				anchor.setColumnIndex(idx + size);
			}
		}		
	}
	private void shiftAfterColumnDelete(int columnIdx, int size) {
		//handling pic shift
		for(PictureAdv pic:pictures){
			NViewAnchor anchor = pic.getAnchor();
			int idx = anchor.getColumnIndex();
			if(idx >= columnIdx+size){
				anchor.setColumnIndex(idx-size);
			}else if(idx >= columnIdx){
				anchor.setColumnIndex(columnIdx);//as excel's rule
				anchor.setXOffset(0);
			}
		}
		//handling pic shift
		for(ChartAdv chart:charts){
			NViewAnchor anchor = chart.getAnchor();
			int idx = anchor.getColumnIndex();
			if(idx >= columnIdx+size){
				anchor.setColumnIndex(idx-size);
			}else if(idx >= columnIdx){
				anchor.setColumnIndex(columnIdx);//as excel's rule
				anchor.setXOffset(0);
			}
		}		
	}
	

	public void deleteRow(int rowIdx, int size) {
		checkOrphan();
		if(size<=0) return;
		
		//clear before move relation
		for(RowAdv row:rows.subValues(rowIdx,rowIdx+size)){
			row.release();
		}		
		rows.delete(rowIdx, size);
		
		shiftAfterRowDelete(rowIdx,size);	
		
		book.sendEvent(ModelEvents.ON_ROW_DELETED, ModelEvents.PARAM_SHEET, this, ModelEvents.PARAM_ROW_INDEX, rowIdx, 
				ModelEvents.PARAM_SIZE, size);
	}
	
	@Override
	void copyTo(SheetAdv sheet) {
		if(sheet==this)
			return;
		
		checkOrphan();
		sheet.checkOrphan();
		if(!getBook().equals(sheet.getBook())){
			throw new UnsupportedOperationException("the source book is different");
		}
		
		
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
		checkOrphan();
		if(size<=0) return;
		
		columns.insert(columnIdx, size);
		
		for(RowAdv row:rows.values()){
			row.insertCell(columnIdx,size);
		}
		
		shiftAfterColumnInsert(columnIdx,size);
		
		book.sendEvent(ModelEvents.ON_COLUMN_INSERTED, ModelEvents.PARAM_SHEET, this, ModelEvents.PARAM_COLUMN_INDEX, columnIdx, 
				ModelEvents.PARAM_SIZE, size);
	}

	public void deleteColumn(int columnIdx, int size) {
		checkOrphan();
		if(size<=0) return;
		
		for(ColumnAdv column:columns.subValues(columnIdx, columnIdx+size)){
			column.release();
		}
		
		columns.delete(columnIdx, size);
		
		for(RowAdv row:rows.values()){
			row.deleteCell(columnIdx,size);
		}
		shiftAfterColumnDelete(columnIdx,size);
		
		book.sendEvent(ModelEvents.ON_COLUMN_DELETED, ModelEvents.PARAM_SHEET, this, ModelEvents.PARAM_COLUMN_INDEX, columnIdx, 
				ModelEvents.PARAM_SIZE, size);
	}

	
	public void checkOrphan(){
		if(book==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
	@Override
	public void release(){
		checkOrphan();
		for(ColumnAdv column:columns.values()){
			column.release();
		}
		for(RowAdv row:rows.values()){
			row.release();
		}
		for(ChartAdv chart:charts){
			chart.release();
		}
		for(PictureAdv picture:pictures){
			picture.release();
		}
		book = null;
		//TODO all 
		
	}

	public String getId() {
		return id;
	}

	public NPicture addPicture(Format format, byte[] data,NViewAnchor anchor) {
		checkOrphan();
		PictureAdv pic = new PictureImpl(this,book.nextObjId("pic"), format, data,anchor);
		pictures.add(pic);
		return pic;
	}
	
	public NPicture getPicture(String picid){
		for(NPicture pic:pictures){
			if(pic.getId().equals(picid)){
				return pic;
			}
		}
		return null;
	}

	public void deletePicture(NPicture picture) {
		checkOrphan();
		checkOwnership(picture);
		((PictureAdv)picture).release();
		pictures.remove(picture);
	}

	public List<NPicture> getPictures() {
		return new ArrayList<NPicture>(pictures);
	}
	
	public NChart addChart(NChart.NChartType type,NViewAnchor anchor) {
		checkOrphan();
		ChartAdv pic = new ChartImpl(this, book.nextObjId("chart"), type, anchor);
		charts.add(pic);
		return pic;
	}
	
	public NChart getChart(String picid){
		for(NChart pic:charts){
			if(pic.getId().equals(picid)){
				return pic;
			}
		}
		return null;
	}

	public void deleteChart(NChart chart) {
		checkOrphan();
		checkOwnership(chart);
		((ChartAdv)chart).release();
		charts.remove(chart);
	}

	public List<NChart> getCharts() {
		return new ArrayList<NChart>(charts);
	}

}
