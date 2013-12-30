/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel.impl;

import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.SortedMap;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.DefaultDataGrid;
import org.zkoss.zss.ngmodel.InvalidateModelOpException;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NChart;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NColumnArray;
import org.zkoss.zss.ngmodel.NDataGrid;
import org.zkoss.zss.ngmodel.NPicture;
import org.zkoss.zss.ngmodel.NPicture.Format;
import org.zkoss.zss.ngmodel.NPrintInfo;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.NViewAnchor;
import org.zkoss.zss.ngmodel.NViewInfo;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class SheetImpl extends AbstractSheetAdv {
	private static final long serialVersionUID = 1L;
	private AbstractBookAdv book;
	private String name;
	private final String id;
	
	private boolean protect;
	
	private NDataGrid dataGrid;
	
	private final IndexPool<AbstractRowAdv> rows = new IndexPool<AbstractRowAdv>(){
		private static final long serialVersionUID = 1L;
		@Override
		void resetIndex(int newidx, AbstractRowAdv obj) {
			obj.setIndex(newidx);
		}};
//	private final BiIndexPool<ColumnAdv> columns = new BiIndexPool<ColumnAdv>();
	private final ColumnArrayPool columnArrays = new ColumnArrayPool();
	
	
	private final List<AbstractPictureAdv> pictures = new LinkedList<AbstractPictureAdv>();
	private final List<AbstractChartAdv> charts = new LinkedList<AbstractChartAdv>();
	
	private final List<CellRegion> mergedRegions = new LinkedList<CellRegion>();
	
	//to store some lowpriority view info
	private final NViewInfo viewInfo = new ViewInfoImpl();
	
	private final NPrintInfo printInfo = new PrintInfoImpl();
	
	private transient HashMap<String,Object> attributes;
	private int defaultColumnWidth = 64; //in pixel
	private int defaultRowHeight = 20;//in pixel
	
	public SheetImpl(AbstractBookAdv book,String id){
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
		return getRow(rowIdx,true);
	}
	@Override
	AbstractRowAdv getRow(int rowIdx, boolean proxy) {
		AbstractRowAdv rowObj = rows.get(rowIdx);
		if(rowObj != null){
			return rowObj;
		}
		return proxy?new RowProxy(this,rowIdx):null;
	}
	@Override
	AbstractRowAdv getOrCreateRow(int rowIdx){
		AbstractRowAdv rowObj = rows.get(rowIdx);
		if(rowObj == null){
			checkOrphan();
			rowObj = new RowImpl(this,rowIdx);
			rows.put(rowIdx, rowObj);
		}
		return rowObj;
	}

	@Override
	public NColumn getColumn(int columnIdx) {
		return getColumn(columnIdx,true);
	}
	
	NColumn getColumn(int columnIdx, boolean proxy) {
		NColumnArray array = getColumnArray(columnIdx);
		if(array==null && !proxy){
			return null;
		}
		return new ColumnProxy(this,columnIdx);
	}
	@Override
	public NColumnArray getColumnArray(int columnIdx) {
		if(columnArrays.hasLastKey(columnIdx)){
			return null;
		}
		SortedMap<Integer, AbstractColumnArrayAdv> submap = columnArrays.lastSubMap(columnIdx);
		
		return submap.size()>0?submap.get(submap.firstKey()):null;
	}
//	@Override
//	ColumnAdv getColumn(int columnIdx, boolean proxy) {
//		ColumnAdv colObj = columns.get(columnIdx);
//		if(colObj != null){
//			return colObj;
//		}
//		return proxy?new ColumnProxy(this,columnIdx):null;
//	}
	
	/**internal use only for developing/test state, should remove when stable*/
	public static boolean DEBUG = false;
	
	private void checkColumnArrayStatus(){
		if(!DEBUG) //only check in dev 
			return;
		
		AbstractColumnArrayAdv prev = null;
		try{
			for(AbstractColumnArrayAdv array:columnArrays.values()){
				//check the existed data
				if(prev==null){
					if(array.getIndex()!=0){
						throw new IllegalStateException("column array doesn't not start with 0 is "+array.getIndex());
					}
				}else{
					if(prev.getLastIndex()+1!=array.getIndex()){
						throw new IllegalStateException("column array doesn't continue, "+prev.getLastIndex() +" to "+array.getIndex());
					}
				}
				prev = array;
			}
		}catch(RuntimeException x){
			System.out.println(">>>>>>>>>>>>>>>>");
			for(AbstractColumnArrayAdv array:columnArrays.values()){
				System.out.println(">>>>"+array.getIndex()+":"+array.getLastIndex());
			}
			System.out.println(">>>>>>>>>>>>>>>>");
			throw x;
		}
		
	}
	

	@Override
	public NColumnArray setupColumnArray(int index, int lastIndex) {
		if(index<0 && lastIndex > index){
			throw new IllegalArgumentException(index+","+lastIndex);
		}
		int start1,end1;
		start1 = end1 = -1;
		
		AbstractColumnArrayAdv ov = columnArrays.overlap(index,lastIndex); 
		if(ov!=null){
			throw new IllegalStateException("Can't setup an overlapped column array "+index+","+lastIndex +" overlppaed "+ov);
		}
		
		
		if(columnArrays.size()==0){
			start1 = 0;
		}else{
			start1 = columnArrays.lastLastKey()+1;
		}
		end1 = index-1;
		
		AbstractColumnArrayAdv array;
		if(start1<=end1 && end1>-1){
			array = new ColumnArrayImpl(this, start1, end1);
			columnArrays.put(array);
		}
		array = new ColumnArrayImpl(this, index, lastIndex);
		columnArrays.put(array);
		
		checkColumnArrayStatus();
		return array;
	}

	
	@Override
	AbstractColumnArrayAdv getOrSplitColumnArray(int columnIdx){
		AbstractColumnArrayAdv contains = (AbstractColumnArrayAdv)getColumnArray(columnIdx);
		if(contains!=null && contains.getIndex()==columnIdx && contains.getLastIndex()==columnIdx){
			return contains;
		}
		
		int start1,end1,start2,end2;
		start1 = end1 = start2 = end2 = -1;
		
		if(contains==null){
			if(columnArrays.size()==0){//no data
				start1 = 0;
			}else{//out of existed array
				start1 = columnArrays.lastLastKey()+1;
			}
			end1 = columnIdx-1;
		}else{
			if(contains.getIndex()==columnIdx){//for the begin
				start2 = columnIdx+1;
				end2 = contains.getLastIndex();
			}else if(contains.getLastIndex()==columnIdx){//at the end
				start1 = contains.getIndex();
				end1 = columnIdx-1;
			}else{
				start1 = contains.getIndex();
				end1 = columnIdx-1;
				end2 = contains.getLastIndex();
				start2 = columnIdx+1;
			}
		}
		AbstractColumnArrayAdv array = null;
		AbstractColumnArrayAdv prev = null;
		if(contains!=null){
			columnArrays.remove(contains);
		}
		//
		if(start2<=end2 && end2>-1){
			prev =new ColumnArrayImpl(this, start2, end2);
			columnArrays.put(prev);
			if(contains!=null){
				prev.setCellStyle(contains.getCellStyle());
				prev.setHidden(contains.isHidden());
				prev.setWidth(contains.getWidth());
			}
		}
		
		array = new ColumnArrayImpl(this, columnIdx, columnIdx);
		columnArrays.put(array);
		if(contains!=null){
			array.setCellStyle(contains.getCellStyle());
			array.setHidden(contains.isHidden());
			array.setWidth(contains.getWidth());
		}
		
		if(start1<=end1 && end1>-1){
			prev =new ColumnArrayImpl(this, start1, end1);
			columnArrays.put(prev);
			if(contains!=null){
				prev.setCellStyle(contains.getCellStyle());
				prev.setHidden(contains.isHidden());
				prev.setWidth(contains.getWidth());
			}
		}
		
		checkColumnArrayStatus();
		return array;
	}
//	@Override
//	int getColumnIndex(ColumnAdv column){
//		return columns.get(column);
//	}

	public NCell getCell(int rowIdx, int columnIdx) {
		return getCell(rowIdx,columnIdx,true);
	}
	
	@Override
	AbstractCellAdv getCell(int rowIdx, int columnIdx, boolean proxy) {
		AbstractRowAdv rowObj = (AbstractRowAdv) getRow(rowIdx,false);
		if(rowObj!=null){
			return rowObj.getCell(columnIdx,proxy);
		}
		return proxy?new CellProxy(this, rowIdx,columnIdx):null;
	}
	@Override
	AbstractCellAdv getOrCreateCell(int rowIdx, int columnIdx){
		AbstractRowAdv rowObj = (AbstractRowAdv)getOrCreateRow(rowIdx);
		AbstractCellAdv cell = rowObj.getOrCreateCell(columnIdx);
		return cell;
	}

	public int getStartRowIndex() {
		int idx1 = rows.firstKey();
		NDataGrid dg = getDataGrid();
		if(dg==null || !dg.isSupportedDataStartEndIndex()){
			return idx1;
		}else{
			int idx2 = dg.getStartRowIndex();
			if(idx1<0)
				return idx2;
			if(idx2<0)
				return idx1;
			return Math.min(idx1, idx2);
		}
	}

	public int getEndRowIndex() {
		int idx1 = rows.lastKey();
		NDataGrid dg = getDataGrid();
		if(dg==null || !dg.isSupportedDataStartEndIndex()){
			return idx1;
		}else{
			int idx2 = dg.getEndRowIndex();
			if(idx1<0)
				return idx2;
			if(idx2<0)
				return idx1;
			return Math.max(idx1, idx2);
		}
	}
	
	public int getStartColumnIndex() {
		return columnArrays.size()>0?columnArrays.firstFirstKey():-1;
	}

	public int getEndColumnIndex() {
		return columnArrays.size()>0?columnArrays.lastLastKey():-1;
	}

	public int getStartCellIndex(int row) {
		int idx1 = -1;
		AbstractRowAdv rowObj = (AbstractRowAdv) getRow(row,false);
		if(rowObj!=null){
			idx1 = rowObj.getStartCellIndex();
		}
		
		NDataGrid dg = getDataGrid();
		if(dg==null || !dg.isSupportedDataStartEndIndex()){
			return idx1;
		}else{
			int idx2 = dg.getStartCellIndex(row);
			if(idx1<0)
				return idx2;
			if(idx2<0)
				return idx1;
			return Math.min(idx1, idx2);
		}
	}

	public int getEndCellIndex(int row) {
		int idx1 = -1;
		AbstractRowAdv rowObj = (AbstractRowAdv) getRow(row,false);
		if(rowObj!=null){
			idx1 = rowObj.getEndCellIndex();
		}
		NDataGrid dg = getDataGrid();
		if(dg==null || !dg.isSupportedDataStartEndIndex()){
			return idx1;
		}else{
			int idx2 = dg.getEndCellIndex(row);
			if(idx1<0)
				return idx2;
			if(idx2<0)
				return idx1;
			return Math.max(idx1, idx2);
		}
	}

	@Override
	void setSheetName(String name) {
		this.name = name;
	}
	@Override
	void onModelInternalEvent(ModelInternalEvent event) {
		for(AbstractRowAdv row:rows.values()){
			row.onModelInternalEvent(event);
		}
		for(AbstractColumnArrayAdv column:columnArrays.values()){
			column.onModelInternalEvent(event);
		}
		//TODO to other object
	}
	
//	public void clearRow(int rowIdx, int rowIdx2) {
//		int start = Math.min(rowIdx, rowIdx2);
//		int end = Math.max(rowIdx, rowIdx2);
//		
//		//clear before move relation
//		for(RowAdv row:rows.subValues(start,end)){
//			row.destroy();
//		}		
//		rows.clear(start,end);
//		
//		//Send event?
//		
//	}

//	public void clearColumn(int columnIdx, int columnIdx2) {
//		int start = Math.min(columnIdx, columnIdx2);
//		int end = Math.max(columnIdx, columnIdx2);
//		
//		
//		for(ColumnAdv column:columns.subValues(start,end)){
//			column.destroy();
//		}
//		columns.clear(start,end);
//		
//		for(RowAdv row:rows.values()){
//			row.clearCell(start,end);
//		}
//		//Send event?
//		
//	}

	public void clearCell(int rowIdx, int columnIdx, int rowIdx2,
			int columnIdx2) {
		int rowStart = Math.min(rowIdx, rowIdx2);
		int rowEnd = Math.max(rowIdx, rowIdx2);
		int columnStart = Math.min(columnIdx, columnIdx2);
		int columnEnd = Math.max(columnIdx, columnIdx2);
		
		Collection<AbstractRowAdv> effected = rows.subValues(rowStart,rowEnd);
		
		Iterator<AbstractRowAdv> iter = effected.iterator();
		while(iter.hasNext()){
			iter.next().clearCell(columnStart, columnEnd);
		}
	}

	public void insertRow(int rowIdx, int size) {
		checkOrphan();
		if(size<=0) return;
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new InvalidateModelOpException("doesn't support insert/delete");
			}
			dg.insertRow(rowIdx, size);
		}
		
		rows.insert(rowIdx, size);
		
		shiftAfterRowInsert(rowIdx,size);
		
		book.sendModelInternalEvent(createModelInternalEvent(ModelInternalEvents.ON_ROW_INSERTED, ModelInternalEvents.PARAM_ROW_INDEX, rowIdx, 
				ModelInternalEvents.PARAM_SIZE, size));
	}
	
	private void shiftAfterRowInsert(int rowIdx, int size) {
		// handling pic shift
		for (AbstractPictureAdv pic : pictures) {
			NViewAnchor anchor = pic.getAnchor();
			int idx = anchor.getRowIndex();
			if (idx >= rowIdx) {
				anchor.setRowIndex(idx + size);
			}
		}
		// handling pic shift
		for (AbstractChartAdv chart : charts) {
			NViewAnchor anchor = chart.getAnchor();
			int idx = anchor.getRowIndex();
			if (idx >= rowIdx) {
				anchor.setRowIndex(idx + size);
			}
		}
	}
	private void shiftAfterRowDelete(int rowIdx, int size) {
		//handling pic shift
		for(AbstractPictureAdv pic:pictures){
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
		for(AbstractChartAdv chart:charts){
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
		for (AbstractPictureAdv pic : pictures) {
			NViewAnchor anchor = pic.getAnchor();
			int idx = anchor.getColumnIndex();
			if (idx >= columnIdx) {
				anchor.setColumnIndex(idx + size);
			}
		}
		// handling pic shift
		for (AbstractChartAdv chart : charts) {
			NViewAnchor anchor = chart.getAnchor();
			int idx = anchor.getColumnIndex();
			if (idx >= columnIdx) {
				anchor.setColumnIndex(idx + size);
			}
		}		
	}
	private void shiftAfterColumnDelete(int columnIdx, int size) {
		//handling pic shift
		for(AbstractPictureAdv pic:pictures){
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
		for(AbstractChartAdv chart:charts){
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
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new InvalidateModelOpException("doesn't support insert/delete");
			}
			dg.deleteRow(rowIdx, size);
		}
		
		//clear before move relation
		for(AbstractRowAdv row:rows.subValues(rowIdx,rowIdx+size)){
			row.destroy();
		}		
		rows.delete(rowIdx, size);
		
		shiftAfterRowDelete(rowIdx,size);	
		
		book.sendModelInternalEvent(createModelInternalEvent(ModelInternalEvents.ON_ROW_DELETED, ModelInternalEvents.PARAM_ROW_INDEX, rowIdx, 
				ModelInternalEvents.PARAM_SIZE, size));
	}
	
	@Override
	void copyTo(AbstractSheetAdv sheet) {
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
		
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new InvalidateModelOpException("doesn't support insert/delete");
			}
			dg.insertColumn(columnIdx, size);
		}
		
		insertAndSplitColumnArray(columnIdx,size);
		
		for(AbstractRowAdv row:rows.values()){
			row.insertCell(columnIdx,size);
		}
		
		shiftAfterColumnInsert(columnIdx,size);
		
		book.sendModelInternalEvent(createModelInternalEvent(ModelInternalEvents.ON_COLUMN_INSERTED, ModelInternalEvents.PARAM_COLUMN_INDEX, columnIdx, 
				ModelInternalEvents.PARAM_SIZE, size));
	}
	
	private void insertAndSplitColumnArray(int columnIdx,int size){
				
		AbstractColumnArrayAdv contains = null;
		
		int start1,end1,start2,end2;
		start1 = end1 = start2 = end2 = -1;
		
		if(columnArrays.hasLastKey(columnIdx)){//no data
			return;
		}
		
		List<AbstractColumnArrayAdv> shift = new LinkedList<AbstractColumnArrayAdv>();
		
		for(AbstractColumnArrayAdv array:columnArrays.lastSubMap(columnIdx).values()){
			if(array.getIndex()<=columnIdx && array.getLastIndex()>=columnIdx){
				contains = array;
			}
			if(array.getIndex()>columnIdx){//shift the right side array
				shift.add(0,array);//revert it to avoid overlap key replace issue
			}
		}
		for(AbstractColumnArrayAdv array:shift){
			columnArrays.remove(array);
			
			array.setIndex(array.getIndex()+size);
			array.setLastIndex(array.getLastIndex()+size);
			
			columnArrays.put(array);
		}
		
		if(contains==null){//doesn't need to do anything
			return;//
		}else{
			if(contains.getIndex()==columnIdx){//from the begin
				start2 = columnIdx+size;
				end2 = contains.getLastIndex()+size;
			}else{//at the end and in the middle
				start1 = contains.getIndex();
				end1 = columnIdx-1;
				start2 = columnIdx+size;
				end2 = contains.getLastIndex()+size;
			}
		}
		
		AbstractColumnArrayAdv array = null;
		AbstractColumnArrayAdv prev = null;
		
		columnArrays.remove(contains);
		//
		if(start2<=end2 && end2>-1){
			prev =new ColumnArrayImpl(this, start2, end2);
			columnArrays.put(prev);
			if(contains!=null){
				prev.setCellStyle(contains.getCellStyle());
				prev.setHidden(contains.isHidden());
				prev.setWidth(contains.getWidth());
			}
		}
		
		array = new ColumnArrayImpl(this, columnIdx, columnIdx+size-1);
		columnArrays.put(array);
		//don't need to copy the property from contains to new inserted array, keep it default.
		
		if(start1<=end1 && end1>-1){
			prev =new ColumnArrayImpl(this, start1, end1);
			columnArrays.put(prev);
			if(contains!=null){
				prev.setCellStyle(contains.getCellStyle());
				prev.setHidden(contains.isHidden());
				prev.setWidth(contains.getWidth());
			}
		}
		
		checkColumnArrayStatus();
	}

	public void deleteColumn(int columnIdx, int size) {
		checkOrphan();
		if(size<=0) return;
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new InvalidateModelOpException("doesn't support insert/delete");
			}
			dg.deleteColumn(columnIdx, size);
		}
		
		deleteAndShrinkColumnArray(columnIdx,size);
		
		for(AbstractRowAdv row:rows.values()){
			row.deleteCell(columnIdx,size);
		}
		shiftAfterColumnDelete(columnIdx,size);
		
		book.sendModelInternalEvent(createModelInternalEvent(ModelInternalEvents.ON_COLUMN_DELETED, ModelInternalEvents.PARAM_COLUMN_INDEX, columnIdx, 
				ModelInternalEvents.PARAM_SIZE, size));
	}
	
	private void deleteAndShrinkColumnArray(int columnIdx,int size){

		if(columnArrays.hasLastKey(columnIdx)){//no data
			return;
		}
		
		List<AbstractColumnArrayAdv> remove = new LinkedList<AbstractColumnArrayAdv>();
		List<AbstractColumnArrayAdv> contains = new LinkedList<AbstractColumnArrayAdv>();
		List<AbstractColumnArrayAdv> leftOver = new LinkedList<AbstractColumnArrayAdv>();
		List<AbstractColumnArrayAdv> rightOver = new LinkedList<AbstractColumnArrayAdv>();
		List<AbstractColumnArrayAdv> right = new LinkedList<AbstractColumnArrayAdv>();
		
		int lastColumnIdx = columnIdx+size-1;
		for(AbstractColumnArrayAdv array:columnArrays.lastSubMap(columnIdx).values()){
			int arrIdx = array.getIndex();
			int arrLastIdx = array.getLastIndex();
			if(arrIdx<columnIdx && arrLastIdx > lastColumnIdx){//array large and contain delete column
				contains.add(array);
			}else if(arrIdx<columnIdx && arrLastIdx >= columnIdx){//overlap left side
				leftOver.add(array);
			}else if(arrIdx >= columnIdx && arrLastIdx <= lastColumnIdx){//contains
				remove.add(array);//remove entire
			}else if(arrIdx<=lastColumnIdx && arrLastIdx > lastColumnIdx){//overlap right side
				rightOver.add(array); 
			}else if(arrIdx>lastColumnIdx){//right side
				right.add(array); 
			}else{
				throw new IllegalStateException("wrong array state");
			}
			
		}
		for(AbstractColumnArrayAdv array:contains){
			columnArrays.remove(array);
			array.setLastIndex(array.getLastIndex()-size);
			columnArrays.put(array);
		}
		for(AbstractColumnArrayAdv array:leftOver){
			columnArrays.remove(array);
			array.setLastIndex(columnIdx-1);//shrink trail
			columnArrays.put(array);
		}
		for(AbstractColumnArrayAdv array:remove){
			columnArrays.remove(array);
		}
		for(AbstractColumnArrayAdv array:rightOver){
			int arrIdx = array.getIndex();
			int arrLastIdx = array.getLastIndex();
			
			columnArrays.remove(array);
			array.setIndex(columnIdx);//shrink head and move trail
			array.setLastIndex(columnIdx + arrLastIdx-lastColumnIdx -1); 
			columnArrays.put(array);
		}
		for(AbstractColumnArrayAdv array:right){
			int arrIdx = array.getIndex();
			int arrLastIdx = array.getLastIndex();
			
			columnArrays.remove(array);
			array.setIndex(arrIdx-size);//shrink head and move trail
			array.setLastIndex(arrLastIdx-size); 
			columnArrays.put(array);
		}	

		checkColumnArrayStatus();
	}

	
	public void checkOrphan(){
		if(book==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
	@Override
	public void destroy(){
		checkOrphan();
		for(AbstractColumnArrayAdv column:columnArrays.values()){
			column.destroy();
		}
		columnArrays.clear();
		for(AbstractRowAdv row:rows.values()){
			row.destroy();
		}
		rows.clear();
		for(AbstractChartAdv chart:charts){
			chart.destroy();
		}
		charts.clear();
		for(AbstractPictureAdv picture:pictures){
			picture.destroy();
		}
		pictures.clear();
		book = null;
		//TODO all 
		
	}

	public String getId() {
		return id;
	}

	public NPicture addPicture(Format format, byte[] data,NViewAnchor anchor) {
		checkOrphan();
		AbstractPictureAdv pic = new PictureImpl(this,book.nextObjId("pic"), format, data,anchor);
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
		((AbstractPictureAdv)picture).destroy();
		pictures.remove(picture);
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public List<NPicture> getPictures() {
		return Collections.unmodifiableList((List)pictures);
	}
	
	public NChart addChart(NChart.NChartType type,NViewAnchor anchor) {
		checkOrphan();
		AbstractChartAdv pic = new ChartImpl(this, book.nextObjId("chart"), type, anchor);
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
		((AbstractChartAdv)chart).destroy();
		charts.remove(chart);
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public List<NChart> getCharts() {
		return Collections.unmodifiableList((List)charts);
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public List<CellRegion> getMergedRegions() {
		return Collections.unmodifiableList((List)mergedRegions);
	}

	@Override
	public void removeMergedRegion(CellRegion region) {
		mergedRegions.remove(region);
	}

	@Override
	public void addMergedRegion(CellRegion region) {
		Validations.argNotNull(region);
		if(region.isSingle()){
			//just ignore it.
		}
		for(CellRegion r:mergedRegions){
			if(r.overlaps(region)){
				throw new InvalidateModelOpException("the region is overlapped "+r+":"+region);
			}
		}
		mergedRegions.add(region);
	}

	@Override
	public List<CellRegion> getOverlappedMergedRegions(CellRegion region) {
		List<CellRegion> list =new LinkedList<CellRegion>(); 
		for(CellRegion r:mergedRegions){
			if(r.overlaps(region)){
				list.add(r);
			}
		}
		return list;
	}

	@Override
	public Object getAttribute(String name) {
		return attributes==null?null:attributes.get(name);
	}

	@Override
	public Object setAttribute(String name, Object value) {
		if(attributes==null){
			attributes = new HashMap<String, Object>();
		}
		return attributes.put(name, value);
	}

	@Override
	public Map<String, Object> getAttributes() {
		return attributes==null?Collections.EMPTY_MAP:Collections.unmodifiableMap(attributes);
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public Iterator<NRow> getRowIterator() {
		NDataGrid dg = getDataGrid();
		if(dg!=null && dg.isSupportedDataIterator()){
			return new JoinRowIterator(this,((Collection)rows.values()).iterator(),dg.getRowIterator());
		}else{
			return Collections.unmodifiableCollection((Collection)rows.values()).iterator();
		}
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public Iterator<NColumnArray> getColumnArrayIterator() {
		return Collections.unmodifiableCollection((Collection)columnArrays.values()).iterator();
	}
	

	@Override
	public Iterator<NColumn> getColumnIterator() {
		return new Iterator<NColumn>(){
			int index = -1;
			@Override
			public boolean hasNext() {
				return getColumnArray(index+1)!=null;
			}

			@Override
			public NColumn next() {
				index++;
				return getColumn(index);
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException("illegal");
			}
		};
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public Iterator<NCell> getCellIterator(int row) {
		
		NDataGrid dg = getDataGrid();
		if(dg!=null && dg.isSupportedDataIterator()){
			return new JoinCellIterator(this,row,(Iterator)((AbstractRowAdv)getRow(row)).getCellIterator(),dg.getCellIterator(row));
		}else{
			return (Iterator)((AbstractRowAdv)getRow(row)).getCellIterator();
		}
	}
	
	@Override
	public int getDefaultRowHeight() {
		return defaultRowHeight;
	}

	@Override
	public int getDefaultColumnWidth() {
		return defaultColumnWidth;
	}

	@Override
	public void setDefaultRowHeight(int height) {
		defaultRowHeight = height;
	}

	@Override
	public void setDefaultColumnWidth(int width) {
		defaultColumnWidth = width;
	}

	@Override
	public int getNumOfPicture() {
		return pictures.size();
	}

	@Override
	public NPicture getPicture(int idx) {
		return pictures.get(idx);
	}

	@Override
	public int getNumOfChart() {
		return charts.size();
	}

	@Override
	public NChart getChart(int idx) {
		return charts.get(idx);
	}

	@Override
	public int getNumOfMergedRegion() {
		return mergedRegions.size();
	}

	@Override
	public CellRegion getMergedRegion(int idx) {
		return mergedRegions.get(idx);
	}

	@Override
	public boolean isProtected() {
		return protect;
	}

	@Override
	public void setProtected(boolean protect) {
		this.protect = protect;
	}

	@Override
	public NViewInfo getViewInfo(){
		return viewInfo;
	}
	
	@Override
	public NPrintInfo getPrintInfo(){
		return printInfo;
	}

	@Override
	public boolean isAutoFilterMode() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public NDataGrid getDataGrid() {
//		if(dataGrid==null){
//			dataGrid = new TreeMapDataGridImpl();
//			dataGrid = new DefaultDataGrid(this);
//		}
		return dataGrid;
	}

	@Override
	public void setDataGrid(NDataGrid dataGrid) {
		this.dataGrid = dataGrid;
	}

}
