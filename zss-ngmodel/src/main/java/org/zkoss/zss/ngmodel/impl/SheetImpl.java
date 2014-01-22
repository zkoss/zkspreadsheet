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

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.SortedMap;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.InvalidateModelOpException;
import org.zkoss.zss.ngmodel.NAutoFilter;
import org.zkoss.zss.ngmodel.NAutoFilter.FilterOp;
import org.zkoss.zss.ngmodel.NAutoFilter.NFilterColumn;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NChart;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NColumnArray;
import org.zkoss.zss.ngmodel.NDataGrid;
import org.zkoss.zss.ngmodel.NDataValidation;
import org.zkoss.zss.ngmodel.NPicture;
import org.zkoss.zss.ngmodel.NPicture.Format;
import org.zkoss.zss.ngmodel.NPrintSetup;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.NSheetViewInfo;
import org.zkoss.zss.ngmodel.NViewAnchor;
import org.zkoss.zss.ngmodel.SheetRegion;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
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
	
	private NAutoFilter autoFilter;
	
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
	private final List<AbstractDataValidationAdv> dataValidations = new LinkedList<AbstractDataValidationAdv>();
	
	private final List<CellRegion> mergedRegions = new LinkedList<CellRegion>();
	
	//to store some lowpriority view info
	private final NSheetViewInfo viewInfo = new SheetViewInfoImpl();
	
	private final NPrintSetup printSetup = new PrintSetupImpl();
	
	private transient HashMap<String,Object> attributes;
	private int defaultColumnWidth = 64; //in pixel
	private int defaultRowHeight = 20;//in pixel
	
	public SheetImpl(AbstractBookAdv book,String id){
		this.book = book;
		this.id = id;
	}
	
	protected void checkOwnership(NPicture picture){
		if(!pictures.contains(picture)){
			throw new IllegalStateException("doesn't has ownership "+ picture);
		}
	}
	
	protected void checkOwnership(NChart chart){
		if(!charts.contains(chart)){
			throw new IllegalStateException("doesn't has ownership "+ chart);
		}
	}
	
	protected void checkOwnership(NDataValidation validation){
		if(!dataValidations.contains(validation)){
			throw new IllegalStateException("doesn't has ownership "+ validation);
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
			if(rowIdx > getBook().getMaxRowIndex()){
				throw new IllegalStateException("can't create the row that exceeds max row size "+getBook().getMaxRowIndex());
			}
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
		
		if(columnIdx > getBook().getMaxColumnIndex()){
			throw new IllegalStateException("can't create the column array that exceeds max row size "+getBook().getMaxRowIndex());
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

	@Override
	public NCell getCell(int rowIdx, int columnIdx) {
		return getCell(rowIdx,columnIdx,true);
	}
	@Override
	public NCell getCell(String cellRef) {
		CellRegion region = new CellRegion(cellRef);
		if(!region.isSingle()){
			throw new InvalidateModelOpException("not a single ref "+cellRef);
		}
		return getCell(region.getRow(),region.getColumn(),true);
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
		return getStartRowIndex(true);
	}

	public int getEndRowIndex() {
		return getEndRowIndex(true);
	}
	@Override
	public int getStartRowIndex(boolean joinDataGrid){
		int idx1 = rows.firstKey();
		NDataGrid dg = getDataGrid();
		if(dg==null || !joinDataGrid){
			return idx1;
		}
		
		int idx2 = dg.getStartRowIndex();
		if(idx1<0){
			return idx2;
		}
		if(idx2<0){
			return idx1;
		}
		return Math.min(idx1, idx2);
	}
	@Override
	public int getEndRowIndex(boolean joinDataGrid){
		int idx1 = rows.lastKey();
		NDataGrid dg = getDataGrid();
		if(dg==null || !joinDataGrid){
			return idx1;
		}
		
		int idx2 = dg.getEndRowIndex();
		if(idx1<0){
			return idx2;
		}
		if(idx2<0){
			return idx1;
		}
		return Math.max(idx1, idx2);
	}
	
	public int getStartColumnIndex() {
		return columnArrays.size()>0?columnArrays.firstFirstKey():-1;
	}

	public int getEndColumnIndex() {
		return columnArrays.size()>0?columnArrays.lastLastKey():-1;
	}

	public int getStartCellIndex(int rowIdx) {
		return getStartCellIndex(rowIdx,true);
	}

	public int getEndCellIndex(int rowIdx) {
		return getEndCellIndex(rowIdx, true);
	}
	@Override
	public int getStartCellIndex(int rowIdx,boolean joinDataGrid){
		int idx1 = -1;
		AbstractRowAdv rowObj = (AbstractRowAdv) getRow(rowIdx,false);
		if(rowObj!=null){
			idx1 = rowObj.getStartCellIndex();
		}
		
		NDataGrid dg = getDataGrid();
		if(dg==null || !joinDataGrid){
			return idx1;
		}
		
		int idx2 = dg.getStartCellIndex(rowIdx);
		if(idx1<0){
			return idx2;
		}
		if(idx2<0){
			return idx1;
		}
		return Math.min(idx1, idx2);
	}
	@Override
	public int getEndCellIndex(int rowIdx,boolean joinDataGrid){
		int idx1 = -1;
		AbstractRowAdv rowObj = (AbstractRowAdv) getRow(rowIdx,false);
		if(rowObj!=null){
			idx1 = rowObj.getEndCellIndex();
		}

		NDataGrid dg = getDataGrid();
		if(dg==null || !joinDataGrid){
			return idx1;
		}
		
		int idx2 = dg.getEndCellIndex(rowIdx);
		if(idx1<0){
			return idx2;
		}
		if(idx2<0){
			return idx1;
		}
		return Math.max(idx1, idx2);
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

	@Override
	public void clearCell(CellRegion region) {
		this.clearCell(region.getRow(), region.getColumn(), region.getLastRow(), region.getLastColumn());
	}
	
	@Override
	public void clearCell(int rowIdx, int columnIdx, int rowIdx2,
			int columnIdx2) {
		int rowStart = Math.min(rowIdx, rowIdx2);
		int rowEnd = Math.max(rowIdx, rowIdx2);
		int columnStart = Math.min(columnIdx, columnIdx2);
		int columnEnd = Math.max(columnIdx, columnIdx2);
		
		Collection<AbstractRowAdv> effected = rows.subValues(rowStart,rowEnd);
		for(AbstractRowAdv row:effected){
			row.clearCell(columnStart, columnEnd);
		}
	}

	@Override
	public void insertRow(int rowIdx, int lastRowIdx) {
		checkOrphan();
		if(rowIdx>lastRowIdx){
			throw new IllegalArgumentException(rowIdx+">"+lastRowIdx);
		}
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new IllegalStateException("doesn't support insert/delete");
			}
			dg.insertRow(rowIdx, lastRowIdx);
		}
		int size = lastRowIdx-rowIdx+1;
		rows.insert(rowIdx, size);

		//destroy the row that exceed the max size
		int maxSize = getBook().getMaxRowSize();
		Collection<AbstractRowAdv> exceeds = new ArrayList<AbstractRowAdv>(rows.subValues(maxSize, Integer.MAX_VALUE));
		if(exceeds.size()>0){
			rows.trim(maxSize);
		}
		for(AbstractRowAdv row:exceeds){
			row.destroy();
		}
		
		shiftAfterRowInsert(rowIdx,size);
		
		book.sendModelInternalEvent(ModelInternalEvents.createModelInternalEvent(ModelInternalEvents.ON_ROW_INSERTED, 
				this, 
				ModelInternalEvents.createDataMap(ModelInternalEvents.PARAM_ROW_INDEX, rowIdx,
						ModelInternalEvents.PARAM_LAST_ROW_INDEX, size)));
	}
	
	@Override
	public void deleteRow(int rowIdx, int lastRowIdx) {
		checkOrphan();
		if(rowIdx>lastRowIdx){
			throw new IllegalArgumentException(rowIdx+">"+lastRowIdx);
		}
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new IllegalStateException("doesn't support insert/delete");
			}
			dg.deleteRow(rowIdx, lastRowIdx);
		}
		
		//clear before move relation
		for(AbstractRowAdv row:rows.subValues(rowIdx,lastRowIdx)){
			row.destroy();
		}
		int size = lastRowIdx-rowIdx+1;
		rows.delete(rowIdx, size);
		
		shiftAfterRowDelete(rowIdx,size);
		
		book.sendModelInternalEvent(ModelInternalEvents.createModelInternalEvent(ModelInternalEvents.ON_ROW_DELETED, 
				this, ModelInternalEvents.createDataMap(ModelInternalEvents.PARAM_ROW_INDEX, rowIdx, 
						ModelInternalEvents.PARAM_ROW_INDEX, lastRowIdx)));
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
		//TODO shift data validation?
		
		NBook book = getBook();
		shiftFormula(new CellRegion(rowIdx,0,book.getMaxRowIndex(),book.getMaxColumnIndex()), size, 0);
		
		
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
		//TODO shift data validation?
		
		shiftFormula(new CellRegion(rowIdx,0,book.getMaxRowIndex(),book.getMaxColumnIndex()), -size, 0);
		
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
		//TODO shift data validation?
		
		shiftFormula(new CellRegion(0,columnIdx,book.getMaxRowIndex(),book.getMaxColumnIndex()), 0, size);
		
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
		//TODO shift data validation?
		
		shiftFormula(new CellRegion(0,columnIdx,book.getMaxRowIndex(),book.getMaxColumnIndex()), 0, -size);
		
	}
	
	@Override
	public void insertCell(CellRegion region,boolean horizontal){
		insertCell(region.getRow(),region.getColumn(),region.getLastRow(),region.getLastColumn(),horizontal);
	}
	@Override
	public void insertCell(int rowIdx,int columnIdx,int lastRowIdx, int lastColumnIdx,boolean horizontal){
		checkOrphan();
		
		if(rowIdx>lastRowIdx){
			throw new IllegalArgumentException(rowIdx+">"+lastRowIdx);
		}
		if(columnIdx>lastColumnIdx){
			throw new IllegalArgumentException(columnIdx+">"+lastColumnIdx);
		}
		
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new IllegalStateException("doesn't support insert/delete");
			}
			//TODO
//			dg.insertCell(rowIdx, columnIdx, lastRowIdx,lastColumnIdx,horizontal);
		}
		
		int columnSize = lastColumnIdx - columnIdx+1;
		int rowSize = lastRowIdx - rowIdx +1; 
		if(horizontal){
			Collection<AbstractRowAdv> effectedRows = rows.subValues(rowIdx,lastRowIdx);
			for(AbstractRowAdv row:effectedRows){
				row.insertCell(columnIdx,columnSize);
			}	
		}else{
			int maxSize = getBook().getMaxRowSize();
			Collection<AbstractRowAdv> effectedRows = rows.descendingSubValues(rowIdx,Integer.MAX_VALUE);
			for(AbstractRowAdv row: new ArrayList<AbstractRowAdv>(effectedRows)){//to aovid concurrent modify
				//move the cell down to target row
				int idx = row.getIndex()+rowSize;
				if(idx >= maxSize){
					//clear the cell since it out of max
					row.clearCell(columnIdx,lastColumnIdx);
				}else{
					AbstractRowAdv target = getOrCreateRow(idx);
					row.moveCellTo(target,columnIdx,lastColumnIdx,0);
				}
			}

		}
		
		shiftAfterCellInsert(rowIdx, columnIdx, lastRowIdx,lastColumnIdx,horizontal);
//		
//		book.sendModelInternalEvent(ModelInternalEvents.createModelInternalEvent(ModelInternalEvents.ON_CELL_INSERTED, 
//				this, 
//				ModelInternalEvents.createDataMap(ModelInternalEvents.PARAM_ROW_INDEX, rowIdx,
//						ModelInternalEvents.PARAM_SIZE, size)));
		
	}

	@Override
	public void deleteCell(CellRegion region,boolean horizontal){
		deleteCell(region.getRow(),region.getColumn(),region.getLastRow(),region.getLastColumn(),horizontal);
	}
	@Override
	public void deleteCell(int rowIdx,int columnIdx,int lastRowIdx, int lastColumnIdx,boolean horizontal){
		checkOrphan();
		if(rowIdx>lastRowIdx){
			throw new IllegalArgumentException(rowIdx+">"+lastRowIdx);
		}
		if(columnIdx>lastColumnIdx){
			throw new IllegalArgumentException(columnIdx+">"+lastColumnIdx);
		}
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new IllegalStateException("doesn't support insert/delete");
			}
			//TODO
//			dg.deleteCell(rowIdx, columnIdx, rowSize,columnSize,horizontal);
		}
		
		int columnSize = lastColumnIdx - columnIdx+1;
		int rowSize = lastRowIdx - rowIdx +1; 
		
		if(horizontal){
			Collection<AbstractRowAdv> effected = rows.subValues(rowIdx,lastRowIdx);
			for(AbstractRowAdv row:effected){
				row.deleteCell(columnIdx,columnSize);
			}	
		}else{
			Collection<AbstractRowAdv> effectedRows = rows.subValues(rowIdx,lastRowIdx);
			for(AbstractRowAdv row:effectedRows){
				row.clearCell(columnIdx,lastColumnIdx);
			}
			effectedRows = rows.subValues(rowIdx+rowSize,Integer.MAX_VALUE);
			for(AbstractRowAdv row: new ArrayList<AbstractRowAdv>(effectedRows)){//to aovid concurrent modify
				//move the cell up
				AbstractRowAdv target = getOrCreateRow(row.getIndex()-rowSize);
				row.moveCellTo(target,columnIdx,lastColumnIdx,0);
			}
		}
		
		shiftAfterCellDelete(rowIdx, columnIdx, lastRowIdx,lastColumnIdx,horizontal);
		
//		book.sendModelInternalEvent(ModelInternalEvents.createModelInternalEvent(ModelInternalEvents.ON_CELL_DELETED, 
//				this, ModelInternalEvents.createDataMap(ModelInternalEvents.PARAM_ROW_INDEX, rowIdx, 
//						ModelInternalEvents.PARAM_SIZE, size)));
	}
	
	
	private void shiftAfterCellInsert(int rowIdx, int columnIdx, int lastRowIdx,
			int lastColumnIdx, boolean horizontal) {
		NBook book = getBook();
		if(horizontal){
			shiftFormula(
				new CellRegion(rowIdx, columnIdx, lastRowIdx, book.getMaxColumnIndex()), 0, lastColumnIdx-columnIdx+1);
		}else{
			shiftFormula(
				new CellRegion(rowIdx, columnIdx, book.getMaxRowIndex(), lastColumnIdx), lastRowIdx - rowIdx + 1, 0);
		}
	}
	private void shiftAfterCellDelete(int rowIdx, int columnIdx, int lastRowIdx,
			int lastColumnIdx, boolean horizontal) {
		NBook book = getBook();
		if(horizontal){
			shiftFormula(
				new CellRegion(rowIdx, columnIdx, lastRowIdx, book.getMaxColumnIndex()), 0, -(lastColumnIdx-columnIdx+1));
		}else{
			shiftFormula(
				new CellRegion(rowIdx, columnIdx, book.getMaxRowIndex(), lastColumnIdx), -(lastRowIdx - rowIdx + 1), 0);
		}
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

	@Override
	public void insertColumn(int columnIdx, int lastColumnIdx) {
		checkOrphan();
		if(columnIdx>lastColumnIdx){
			throw new IllegalArgumentException(columnIdx+">"+lastColumnIdx);
		}
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new IllegalStateException("doesn't support insert/delete");
			}
			dg.insertColumn(columnIdx, lastColumnIdx);
		}
		int size = lastColumnIdx - columnIdx + 1;
		insertAndSplitColumnArray(columnIdx,size);
		
		for(AbstractRowAdv row:rows.values()){
			row.insertCell(columnIdx,size);
		}
		
		shiftAfterColumnInsert(columnIdx,size);
		
		book.sendModelInternalEvent(ModelInternalEvents.createModelInternalEvent(ModelInternalEvents.ON_COLUMN_INSERTED, this,
				ModelInternalEvents.createDataMap(ModelInternalEvents.PARAM_COLUMN_INDEX, columnIdx, 
						ModelInternalEvents.PARAM_LAST_COLUMN_INDEX, lastColumnIdx)));
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
		
		//destroy the cell that exceeds the max size
		int maxSize = getBook().getMaxColumnSize();
		Collection<AbstractColumnArrayAdv> exceeds = new ArrayList<AbstractColumnArrayAdv>(columnArrays.firstSubValues(maxSize, Integer.MAX_VALUE));
		if(exceeds.size()>0){
			columnArrays.trim(maxSize);
		}
		for(AbstractColumnArrayAdv ca:exceeds){
			ca.destroy();
		}
		
		
		checkColumnArrayStatus();
	}

	@Override
	public void deleteColumn(int columnIdx, int lastColumnIdx) {
		checkOrphan();
		if(columnIdx>lastColumnIdx){
			throw new IllegalArgumentException(columnIdx+">"+lastColumnIdx);
		}
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new IllegalStateException("doesn't support insert/delete");
			}
			dg.deleteColumn(columnIdx, lastColumnIdx);
		}
		int size = lastColumnIdx - columnIdx + 1;
		deleteAndShrinkColumnArray(columnIdx,size);
		
		for(AbstractRowAdv row:rows.values()){
			row.deleteCell(columnIdx,size);
		}
		shiftAfterColumnDelete(columnIdx,size);
		
		book.sendModelInternalEvent(ModelInternalEvents.createModelInternalEvent(ModelInternalEvents.ON_COLUMN_DELETED, 
				this, ModelInternalEvents.createDataMap(ModelInternalEvents.PARAM_COLUMN_INDEX, columnIdx, 
						ModelInternalEvents.PARAM_LAST_COLUMN_INDEX, lastColumnIdx)));
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
	@Override
	public void moveCell(CellRegion region,int rowOffset, int columnOffset) {
		this.moveCell(region.getRow(), region.getColumn(), region.getLastRow(), region.getLastColumn(),rowOffset,columnOffset);
	}
	
	@Override
	public void moveCell(int rowIdx, int columnIdx,int lastRowIdx,int lastColumnIdx, int rowOffset, int columnOffset){
		if(rowOffset==0 && columnOffset==0)
			return;
		
		int maxRow = getBook().getMaxRowIndex();
		int maxCol = getBook().getMaxColumnIndex();
		
		if(rowIdx<0 || columnIdx<0 || 
				rowIdx > lastRowIdx || lastRowIdx > maxRow || columnIdx>lastColumnIdx || lastColumnIdx>maxCol){
			throw new InvalidateModelOpException(new CellRegion(rowIdx,columnIdx,lastRowIdx,lastColumnIdx).getReferenceString()+" is illegal");
		}
		
		if(rowIdx+rowOffset<0 || columnIdx+columnOffset<0 || 
				lastRowIdx+rowOffset > maxRow|| lastColumnIdx+columnOffset > maxCol){
			throw new InvalidateModelOpException(new CellRegion(rowIdx,columnIdx,lastRowIdx,lastColumnIdx).getReferenceString()+" can't move to offset "+rowOffset+","+columnOffset);
		}
		
		//TODO optimal for whole row, whole column
		
		NDataGrid dg = getDataGrid();
		if(dg!=null){
			if(!dg.isSupportedOperations()){
				throw new InvalidateModelOpException("doesn't support insert/delete");
			}
			//TODO
//			dg.moveCell(rowIdx, columnIdx, rowSize,columnSize,horizontal);
		}
		
		//check merge overlaps and contains
		CellRegion sreRegion = new CellRegion(rowIdx,columnIdx,lastRowIdx,lastColumnIdx);
		Collection<CellRegion> containsMerge = getContainsMergedRegions(sreRegion);
		Collection<CellRegion> overlapsMerge = getOverlapsMergedRegions(sreRegion);
		if(containsMerge.size()!=overlapsMerge.size()){
			ArrayList<CellRegion> ov = new ArrayList<CellRegion>(overlapsMerge);
			ov.removeAll(containsMerge);
			throw new InvalidateModelOpException("can't move "+sreRegion.getReferenceString()+" which overlaps merge area "+ov.get(0).getReferenceString());
		}
		CellRegion targetRegion = new CellRegion(rowIdx+rowOffset,columnIdx+columnOffset,lastRowIdx+rowOffset,lastColumnIdx+columnOffset);
		if(getOverlapsMergedRegions(targetRegion).size()>0){
			throw new InvalidateModelOpException("can't move to "+targetRegion.getReferenceString()+" which overlaps merge area");
		}
		
		
		boolean reverseYDir = rowOffset>0;
		boolean reverseXDir = columnOffset>0;
		
		int rowStart = reverseYDir?lastRowIdx:rowIdx;
		int rowEnd = reverseYDir?rowIdx:lastRowIdx;
		int colStart = reverseXDir?lastColumnIdx:columnIdx;
		int colEnd = reverseXDir?columnIdx:lastColumnIdx;
		
		for(int r = rowStart; reverseYDir?r>=rowEnd:r<=rowEnd;){
			int tr = r+rowOffset;
			AbstractRowAdv row = getRow(r,false);
			for(int c = colStart; reverseXDir?c>=colEnd:c<=colEnd;){
				int tc = c+columnOffset;
				NCell cell = row==null?null:row.getCell(c, false);
				if(cell==null){ // no such cell, the clear the target cell
					clearCell(tr, tc, tr, tc);
				}else{
					AbstractRowAdv target = getOrCreateRow(tr);
					row.moveCellTo(target, c, c, columnOffset);
				}
				
				if(reverseXDir){
					c--;
				}else{
					c++;
				}				
			}
			if(reverseYDir){
				r--;
			}else{ 
				r++;
			}
		}
		ModelUpdateUtil.addCellUpdate(rowIdx,columnIdx,lastRowIdx,lastColumnIdx);
		ModelUpdateUtil.addCellUpdate(rowIdx+rowOffset,columnIdx+columnOffset,lastRowIdx+rowOffset,lastColumnIdx+columnOffset);
		//shift the merge
		mergedRegions.removeAll(containsMerge);
		for(CellRegion merge:containsMerge){
			CellRegion newMerge = new CellRegion(merge.getRow() + rowOffset,merge.getColumn()+ columnOffset,
					merge.getLastRow()+rowOffset,merge.getLastColumn()+columnOffset);
			mergedRegions.add(newMerge);
			ModelUpdateUtil.addMergeUpdate(merge, newMerge);
		}
		
		shiftAfterCellMove(rowIdx, columnIdx,lastRowIdx,lastColumnIdx, rowOffset, columnOffset);
		
		//TODO validation and other stuff
		
		//TODO event
	}

	
	private void shiftAfterCellMove(int rowIdx, int columnIdx, int lastRowIdx,
			int lastColumnIdx, int rowOffset, int columnOffset) {
		shiftFormula(new CellRegion(rowIdx,columnIdx,lastRowIdx,lastColumnIdx),rowOffset,columnOffset);
	}
	
	private void shiftFormula(CellRegion src,int rowOffset,int columnOffset){
		NBook book = getBook();
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)book.getBookSeries();
		DependencyTable dt = bs.getDependencyTable();
		Set<Ref> dependents = dt.getDependents(new RefImpl(book.getBookName(),getSheetName(),src.getRow(),src.getColumn(),src.getLastRow(),src.getLastColumn()));
		FormulaShiftHelper shiftHelper = new FormulaShiftHelper(bs,new SheetRegion(this,src),rowOffset,columnOffset);
		shiftHelper.shift(dependents);
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
		for(AbstractDataValidationAdv validation:dataValidations){
			validation.destroy();
		}
		dataValidations.clear();
		
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
	@Override
	public List<NPicture> getPictures() {
		return Collections.unmodifiableList((List)pictures);
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
	public NChart addChart(NChart.NChartType type,NViewAnchor anchor) {
		checkOrphan();
		AbstractChartAdv pic = new ChartImpl(this, book.nextObjId("chart"), type, anchor);
		charts.add(pic);
		return pic;
	}
	@Override
	public NChart getChart(String picid){
		for(NChart pic:charts){
			if(pic.getId().equals(picid)){
				return pic;
			}
		}
		return null;
	}
	@Override
	public void deleteChart(NChart chart) {
		checkOrphan();
		checkOwnership(chart);
		((AbstractChartAdv)chart).destroy();
		charts.remove(chart);
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	@Override
	public List<NChart> getCharts() {
		return Collections.unmodifiableList((List)charts);
	}

	@Override
	public int getNumOfChart() {
		return charts.size();
	}

	@Override
	public NChart getChart(int idx) {
		return charts.get(idx);
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public List<CellRegion> getMergedRegions() {
		return Collections.unmodifiableList((List)mergedRegions);
	}

	@Override
	public void removeMergedRegion(CellRegion region) {
		mergedRegions.remove(region);
		ModelUpdateUtil.addMergeUpdate(region, null);
	}

	@Override
	public void addMergedRegion(CellRegion region) {
		Validations.argNotNull(region);
		if(region.isSingle()){
			return;
		}
		for(CellRegion r:mergedRegions){
			if(r.overlaps(region)){
				throw new InvalidateModelOpException("the region is overlapped "+r+":"+region);
			}
		}
		mergedRegions.add(region);
		ModelUpdateUtil.addMergeUpdate(null, region);
	}

	@Override
	public List<CellRegion> getOverlapsMergedRegions(CellRegion region) {
		List<CellRegion> list =new LinkedList<CellRegion>(); 
		for(CellRegion r:mergedRegions){
			if(r.overlaps(region)){
				list.add(r);
			}
		}
		return list;
	}
	@Override
	public List<CellRegion> getContainsMergedRegions(CellRegion region) {
		List<CellRegion> list =new LinkedList<CellRegion>(); 
		for(CellRegion r:mergedRegions){
			if(region.contains(r)){
				list.add(r);
			}
		}
		return list;
	}	
	
	@Override
	public CellRegion getMergedRegion(String cellRef) {
		CellRegion region = new CellRegion(cellRef);
		if(!region.isSingle()){
			throw new InvalidateModelOpException("not a single ref "+cellRef);
		}
		return getMergedRegion(region.getRow(),region.getColumn());
	}
	@Override
	public CellRegion getMergedRegion(int row, int column) {
		List<CellRegion> list =new LinkedList<CellRegion>(); 
		for(CellRegion r:mergedRegions){
			if(r.contains(row, column)){
				return r;
			}
		}
		return null;
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

	@Override
	public Iterator<NRow> getRowIterator() {
		return getRowIterator(true);
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public Iterator<NRow> getRowIterator(boolean joinDataGrid) {
		NDataGrid dg = getDataGrid();
		if(joinDataGrid && dg!=null && dg.isProvidedIterator()){
			return new JoinRowIterator(this,Collections.unmodifiableCollection((Collection)rows.values()).iterator(),dg.getRowIterator());
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
				throw new UnsupportedOperationException("readonly");
			}
		};
	}
	
	@Override
	public Iterator<NCell> getCellIterator(int row) {
		return getCellIterator(row,true);
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public Iterator<NCell> getCellIterator(int row,boolean joinDataGrid) {
		NDataGrid dg = getDataGrid();
		if(joinDataGrid && dg!=null && dg.isProvidedIterator()){
			return new JoinCellIterator(this,row,(Iterator)((AbstractRowAdv)getRow(row)).getCellIterator(false),dg.getCellIterator(row));
		}else{
			return (Iterator)((AbstractRowAdv)getRow(row)).getCellIterator(false);
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
	public NSheetViewInfo getViewInfo(){
		return viewInfo;
	}
	
	@Override
	public NPrintSetup getPrintSetup(){
		return printSetup;
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
	
	@Override
	public NDataValidation addDataValidation(CellRegion region) {
		checkOrphan();
		AbstractDataValidationAdv validation = new DataValidationImpl(this, book.nextObjId("valid"));
		validation.addRegion(region);
		dataValidations.add(validation);
		return validation;
	}
	@Override
	public NDataValidation getDataValidation(String validationid){
		for(NDataValidation validation:dataValidations){
			if(validation.getId().equals(validationid)){
				return validation;
			}
		}
		return null;
	}
	@Override
	public void deleteDataValidation(NDataValidation validationid) {
		checkOrphan();
		checkOwnership(validationid);
		((AbstractDataValidationAdv)validationid).destroy();
		dataValidations.remove(validationid);
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public List<NDataValidation> getDataValidations() {
		return Collections.unmodifiableList((List)dataValidations);
	}

	@Override
	public int getNumOfDataValidation() {
		return dataValidations.size();
	}

	@Override
	public NDataValidation getDataValidation(int idx) {
		return dataValidations.get(idx);
	}
	
	@Override
	public NDataValidation getDataValidation(int rowIdx,int columnIdx) {
		for(NDataValidation validation:dataValidations){
			for(CellRegion region:validation.getRegions()){
				if(region.contains(rowIdx, columnIdx)){
					return validation;
				}
			}
		}
		return null;
	}

	@Override
	public NAutoFilter getAutoFilter() {
		return autoFilter;
	}

	@Override
	public NAutoFilter createAutoFilter(CellRegion region) {
		Validations.argNotNull(region);
		
		autoFilter = new AutoFilterImpl(region);
		final int left = region.getColumn();
        final int top = region.getRow();
        final int right = region.getLastColumn();
        final int bottom = region.getLastRow();
        
		//refer from XSSFSheet impl
		//handle the showButton on merged cell
		for (CellRegion mrng:getMergedRegions()) {
			final int t = mrng.getRow();
	        final int b = mrng.getLastRow();
	        final int l = mrng.getColumn();
	        final int r = mrng.getLastColumn();
	        
	        if (t == top && l <= right && l >= left) { // to be add filter column to hide button
	        	for(int c = l; c < r; ++c) {
		        	final int colId = c - left; 
		        	final NFilterColumn col = autoFilter.getFilterColumn(colId, true);
		        	col.setProperties(FilterOp.AND, null, null, false);
	        	}
	        }
		}
		
		
		return autoFilter;
	}
	
	@Override
	public void deleteAutoFilter() {
		autoFilter = null;
	}

	@Override
	public void clearAutoFilter() {
		autoFilter = null;
	}

}
