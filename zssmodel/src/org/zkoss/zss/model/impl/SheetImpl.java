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
package org.zkoss.zss.model.impl;

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

import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.poi.ss.util.WorkbookUtil;
import org.zkoss.util.logging.Log;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.InvalidateModelOpException;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SChart;
import org.zkoss.zss.model.SColumn;
import org.zkoss.zss.model.SColumnArray;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SPicture;
import org.zkoss.zss.model.SPrintSetup;
import org.zkoss.zss.model.SRow;
import org.zkoss.zss.model.SSheetViewInfo;
import org.zkoss.zss.model.ViewAnchor;
import org.zkoss.zss.model.PasteOption;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.model.SPicture.Format;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class SheetImpl extends AbstractSheetAdv {
	private static final long serialVersionUID = 1L;
	private static final Log logger = Log.lookup(SheetImpl.class);
			
	private AbstractBookAdv book;
	private String name;
	private final String id;
	
	private String password;
	
	private SAutoFilter autoFilter;
	
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
	private final SSheetViewInfo viewInfo = new SheetViewInfoImpl();
	
	private final SPrintSetup printSetup = new PrintSetupImpl();
	
	private transient HashMap<String,Object> attributes;
	private int defaultColumnWidth = 64; //in pixel
	private int defaultRowHeight = 20;//in pixel
	
	public SheetImpl(AbstractBookAdv book,String id){
		this.book = book;
		this.id = id;
	}
	
	protected void checkOwnership(SPicture picture){
		if(!pictures.contains(picture)){
			throw new IllegalStateException("doesn't has ownership "+ picture);
		}
	}
	
	protected void checkOwnership(SChart chart){
		if(!charts.contains(chart)){
			throw new IllegalStateException("doesn't has ownership "+ chart);
		}
	}
	
	protected void checkOwnership(SDataValidation validation){
		if(!dataValidations.contains(validation)){
			throw new IllegalStateException("doesn't has ownership "+ validation);
		}
	}
	
	public SBook getBook() {
		checkOrphan();
		return book;
	}

	public String getSheetName() {
		return name;
	}

	public SRow getRow(int rowIdx) {
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
	public SColumn getColumn(int columnIdx) {
		return getColumn(columnIdx,true);
	}
	
	SColumn getColumn(int columnIdx, boolean proxy) {
		SColumnArray array = getColumnArray(columnIdx);
		if(array==null && !proxy){
			return null;
		}
		return new ColumnProxy(this,columnIdx);
	}
	@Override
	public SColumnArray getColumnArray(int columnIdx) {
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
			if(logger.debugable()){
				for(AbstractColumnArrayAdv array:columnArrays.values()){
					logger.debug("ColumnArray "+array.getIndex()+":"+array.getLastIndex());
				}
			}
			throw x;
		}
		
	}
	

	@Override
	public SColumnArray setupColumnArray(int index, int lastIndex) {
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
	public SCell getCell(int rowIdx, int columnIdx) {
		return getCell(rowIdx,columnIdx,true);
	}
	@Override
	public SCell getCell(String cellRef) {
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
		return rows.firstKey();
	}

	public int getEndRowIndex() {
		return rows.lastKey();
	}
	
	public int getStartColumnIndex() {
		return columnArrays.size()>0?columnArrays.firstFirstKey():-1;
	}

	public int getEndColumnIndex() {
		return columnArrays.size()>0?columnArrays.lastLastKey():-1;
	}

	public int getStartCellIndex(int rowIdx) {
		int idx1 = -1;
		AbstractRowAdv rowObj = (AbstractRowAdv) getRow(rowIdx,false);
		if(rowObj!=null){
			idx1 = rowObj.getStartCellIndex();
		}
		return idx1;

	}

	public int getEndCellIndex(int rowIdx) {
		int idx1 = -1;
		AbstractRowAdv rowObj = (AbstractRowAdv) getRow(rowIdx,false);
		if(rowObj!=null){
			idx1 = rowObj.getEndCellIndex();
		}
		return idx1;
	}
	
	@Override
	void setSheetName(String name) {
		checkLegalSheetName(name);
		this.name = name;
	}
	

	private void checkLegalSheetName(String name) {
		try{
			WorkbookUtil.validateSheetName(name);
		}catch(IllegalArgumentException x){
			throw new InvalidateModelOpException(x.getMessage());
		}catch(Exception x){
			throw new InvalidateModelOpException("The sheet name "+name+" is not allowed");
		}
	}
//	@Override
//	void onModelInternalEvent(ModelInternalEvent event) {
//		for(AbstractRowAdv row:rows.values()){
//			row.onModelInternalEvent(event);
//		}
//		for(AbstractColumnArrayAdv column:columnArrays.values()){
//			column.onModelInternalEvent(event);
//		}
//		//TODO to other object
//	}
	
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
		Map<String,Object> dataBefore = shiftBeforeRowInsert(rowIdx,lastRowIdx);
		ModelUpdateUtil.addInsertDeleteUpdate(this, true, true, rowIdx, lastRowIdx);
		shiftAfterRowInsert(dataBefore,rowIdx,lastRowIdx);
	}
	
	private Map<String, Object> shiftBeforeRowInsert(int rowIdx, int lastRowIdx) {
		Map<String,Object> dataBefore = new HashMap<String, Object>();
		// handling merged regions shift
		// find merge cells in affected region and remove them
		CellRegion affectedRegion = new CellRegion(rowIdx, 0, book.getMaxRowIndex(), book.getMaxColumnIndex());
		List<CellRegion> toExtend = getOverlapsMergedRegions(affectedRegion, true);
		List<CellRegion> toShift = getContainsMergedRegions(affectedRegion);
		removeMergedRegion(affectedRegion, true);
		dataBefore.put("toExtend", toExtend);
		dataBefore.put("toShift", toShift);
		return dataBefore;
	}

	private void shiftAfterRowInsert(Map<String, Object> dataBefore, int rowIdx, int lastRowIdx) {
		// handling pic location shift
		int size = lastRowIdx - rowIdx+1;
		for (AbstractPictureAdv pic : pictures) {
			ViewAnchor anchor = pic.getAnchor();
			int idx = anchor.getRowIndex();
			if (idx >= rowIdx) {
				anchor.setRowIndex(idx + size);
			}
		}
		// handling chart location shift
		for (AbstractChartAdv chart : charts) {
			ViewAnchor anchor = chart.getAnchor();
			int idx = anchor.getRowIndex();
			if (idx >= rowIdx) {
				anchor.setRowIndex(idx + size);
			}
		}
		
		// handling merged regions shift
		List<CellRegion> toExtend = (List<CellRegion>)dataBefore.get("toExtend");
		List<CellRegion> toShift = (List<CellRegion>)dataBefore.get("toShift");
		// extend/move removed merged cells, then add them back
		for(CellRegion r : toExtend) {
			addMergedRegion(new CellRegion(r.row, r.column, r.lastRow + size, r.lastColumn));
		}
		for(CellRegion r : toShift) {
			addMergedRegion(new CellRegion(r.row + size, r.column, r.lastRow + size, r.lastColumn));
		}
		
		//TODO shift data validation?
		for(AbstractDataValidationAdv validation:dataValidations){
			//TODO zss 3.5
//			int idx =0;
//			Set<CellRegion> remove = new HashSet<CellRegion>();
//			boolean dirty = false;
//			for(CellRegion region:validation.getRegions()){
//				//TODO
//				CellRegion shifted = null; //TODO Shift
//				if(shifted == null){
//					remove.add(region);
//				}else if(!shifted.equals(region)){
//					validation.setRegion(idx, region);
//					dirty = true;
//				}
//				idx++;
//			}
//			for(CellRegion r:remove){
//				validation.removeRegion(r);
//				dirty = true;
//			}
//			
//			if(dirty){
//				((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_DATA_VALIDATION_CONTENT_CHANGE,this,
//					ModelEvents.createDataMap(ModelEvents.PARAM_OBJECT_ID,validation.getId())));
//			}
			
		}
		
		extendFormula(new CellRegion(rowIdx,0,lastRowIdx,book.getMaxColumnIndex()), false);
	}	
	
	@Override
	public void deleteRow(int rowIdx, int lastRowIdx) {
		checkOrphan();
		if(rowIdx>lastRowIdx){
			throw new IllegalArgumentException(rowIdx+">"+lastRowIdx);
		}
		
		//clear before move relation
		for(AbstractRowAdv row:rows.subValues(rowIdx,lastRowIdx)){
			row.destroy();
		}
		int size = lastRowIdx-rowIdx+1;
		rows.delete(rowIdx, size);
		Map<String,Object> dataBefore = shiftBeforeRowDelete(rowIdx,lastRowIdx);
		ModelUpdateUtil.addInsertDeleteUpdate(this, false, true, rowIdx, lastRowIdx);
		shiftAfterRowDelete(dataBefore,rowIdx,lastRowIdx);
	}	
	
	private Map<String, Object> shiftBeforeRowDelete(int rowIdx, int lastRowIdx) {
		Map<String,Object> dataBefore = new HashMap<String,Object>();
		CellRegion affectedRegion = new CellRegion(rowIdx, 0, book.getMaxRowIndex(), book.getMaxColumnIndex());
		List<CellRegion> toShrink = getOverlapsMergedRegions(affectedRegion, false);
		removeMergedRegion(affectedRegion, true);
		dataBefore.put("toShrink", toShrink);
		return dataBefore;
	}

	private void shiftAfterRowDelete(Map<String, Object> dataBefore, int rowIdx, int lastRowIdx) {
		//handling pic shift
		int size = lastRowIdx - rowIdx+1;
		for(AbstractPictureAdv pic:pictures){
			ViewAnchor anchor = pic.getAnchor();
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
			ViewAnchor anchor = chart.getAnchor();
			int idx = anchor.getRowIndex();
			if(idx >= rowIdx+size){
				anchor.setRowIndex(idx-size);
			}else if(idx >= rowIdx){
				anchor.setRowIndex(rowIdx);//as excel's rule
				anchor.setYOffset(0);
			}
		}
		
		// handling merged regions shift
		List<CellRegion> toShrink = (List<CellRegion>)dataBefore.get("toShrink");
		// shrink/move removed merged cells, then add them back
		for(CellRegion r : toShrink) {
			CellRegion shrank = shrinkRow(r, rowIdx, lastRowIdx);
			if(shrank != null) {
				addMergedRegion(shrank);
			}
		}

		//TODO shift data validation?
		
		shrinkFormula(new CellRegion(rowIdx,0,lastRowIdx,book.getMaxColumnIndex()), false);
		
	}
	
	private CellRegion shrinkRow(CellRegion region, int row, int lastRow) {
		int[] shrank = shrinkIndexes(region.row, region.lastRow, row, lastRow);
		return shrank != null ? new CellRegion(shrank[0], region.column, shrank[1], region.lastColumn) : null;
	}
	
	private CellRegion shrinkColumn(CellRegion region, int column, int lastColumn) {
		int[] shrank = shrinkIndexes(region.column, region.lastColumn, column, lastColumn);
		return shrank != null ? new CellRegion(region.row, shrank[0], region.lastRow, shrank[1]) : null;
	}
	
	/**
	 * To shrink (and shift) line ab by given line cd, a~d are indexes.
	 * @return a shrank line, or null if deleted completely.
	 */
	private int[] shrinkIndexes(int a, int b, int c, int d) {
		if(a < 0 || b < 0 || c < 0 || d < 0 || a > b || c > d) {
			throw new IllegalArgumentException("indexes must be >= 0 and ascending");
		}

		// check every point at shrinking segment or after it
		// then move to corresponding position,
		// e.g b point at shrinking segment, so move it to c-1:
		//        c       d               c       d
		// ---+---+-x-+-x-+--- >>> ---+---+-x-+-x-+---  
		//    a       b               a  b 
		int delta = d - c + 1;
		if(a > d) {
			a -= delta;
		} else if(a >= c) {
			a = c; // as (d + 1) - delta
			a = (d + 1) - delta;
		}
		if(b > d) {
			b -= delta;
		} else if(b >= c) {
			b = c - 1; //
		}
		return (b >= a) ? new int[]{a, b} : null;
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
		
		int columnSize = lastColumnIdx - columnIdx+1;
		int rowSize = lastRowIdx - rowIdx +1; 
		if(horizontal){
			Collection<AbstractRowAdv> effectedRows = rows.subValues(rowIdx,lastRowIdx);
			for(AbstractRowAdv row:effectedRows){
				row.insertCell(columnIdx,columnSize);
			}
			
			// notify affected region update
			ModelUpdateUtil.addCellUpdate(this, rowIdx, columnIdx, lastRowIdx, getBook().getMaxColumnIndex());
		}else{ // vertical
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
			
			// notify affected region update
			ModelUpdateUtil.addCellUpdate(this, rowIdx, columnIdx, getBook().getMaxRowIndex(), lastColumnIdx);
		}
		
		shiftAfterCellInsert(rowIdx, columnIdx, lastRowIdx,lastColumnIdx,horizontal);
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
		
		int columnSize = lastColumnIdx - columnIdx+1;
		int rowSize = lastRowIdx - rowIdx +1; 
		
		if(horizontal){
			Collection<AbstractRowAdv> effected = rows.subValues(rowIdx,lastRowIdx);
			for(AbstractRowAdv row:effected){
				row.deleteCell(columnIdx,columnSize);
			}
			
			// notify affected region update
			ModelUpdateUtil.addCellUpdate(this, rowIdx, columnIdx, lastRowIdx, getBook().getMaxColumnIndex());
		}else{ // vertical
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
			
			// notify affected region update
			ModelUpdateUtil.addCellUpdate(this, rowIdx, columnIdx, getBook().getMaxRowIndex(), lastColumnIdx);
		}
		
		shiftAfterCellDelete(rowIdx, columnIdx, lastRowIdx,lastColumnIdx,horizontal);
	}
	
	
	private void shiftAfterCellInsert(int rowIdx, int columnIdx, int lastRowIdx,
			int lastColumnIdx, boolean horizontal) {
		
		// handle merged cells
		// move merged cells only if they are contained in affected region
		// and unmerge others if they are overlapped with affected region 
		if(horizontal) {
			// find merge cells in affected region and remove them
			int size = lastColumnIdx - columnIdx + 1;
			CellRegion affectedRegion = new CellRegion(rowIdx, columnIdx, lastRowIdx, book.getMaxColumnIndex());
			List<CellRegion> toShift = getContainsMergedRegions(affectedRegion);
			removeMergedRegion(affectedRegion, true); // including contained and overlapped
			for(CellRegion r : toShift) { // only add contained back, to simulate shifting
				addMergedRegion(new CellRegion(r.row, r.column + size, r.lastRow, r.lastColumn + size));
			}

		} else { // vertical
			
			// find merge cells in affected region and remove them
			int size = lastRowIdx - rowIdx + 1;
			CellRegion affectedRegion = new CellRegion(rowIdx, columnIdx, book.getMaxRowIndex(), lastColumnIdx);
			List<CellRegion> toShift = getContainsMergedRegions(affectedRegion);
			removeMergedRegion(affectedRegion, true); // including contained and overlapped
			for(CellRegion r : toShift) { // only add contained back, to simulate shifting
				addMergedRegion(new CellRegion(r.row + size, r.column, r.lastRow + size, r.lastColumn));
			}
		}
		
		extendFormula(new CellRegion(rowIdx, columnIdx, lastRowIdx, lastColumnIdx),horizontal);
	}
	
	private void shiftAfterCellDelete(int rowIdx, int columnIdx, int lastRowIdx,
			int lastColumnIdx, boolean horizontal) {
		
		// handle merged cells
		// unmerge every merged cells overlapped with delete region
		removeMergedRegion(new CellRegion(rowIdx, columnIdx, lastRowIdx, lastColumnIdx), true);
		// move merged cells only if they are contained in affected region
		// and unmerge others if they are overlapped with affected region 
		if(horizontal) {
			// find merge cells in affected region and remove them
			int size = lastColumnIdx - columnIdx + 1;
			CellRegion affectedRegion = new CellRegion(rowIdx, lastColumnIdx, lastRowIdx, book.getMaxColumnIndex());
			List<CellRegion> toShift = getContainsMergedRegions(affectedRegion);
			removeMergedRegion(affectedRegion, true); // including contained and overlapped
			for(CellRegion r : toShift) { // only add contained back, to simulate shifting
				addMergedRegion(new CellRegion(r.row, r.column - size, r.lastRow, r.lastColumn - size));
			}

		} else { // vertical
			
			// find merge cells in affected region and remove them
			int size = lastRowIdx - rowIdx + 1;
			CellRegion affectedRegion = new CellRegion(lastRowIdx, columnIdx, book.getMaxRowIndex(), lastColumnIdx);
			List<CellRegion> toShift = getContainsMergedRegions(affectedRegion);
			removeMergedRegion(affectedRegion, true); // including contained and overlapped
			for(CellRegion r : toShift) { // only add contained back, to simulate shifting
				addMergedRegion(new CellRegion(r.row - size, r.column, r.lastRow - size, r.lastColumn));
			}
		}
		
		shrinkFormula(new CellRegion(rowIdx, columnIdx, lastRowIdx, lastColumnIdx),horizontal);
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
		builder.append("  ==Row=={");
		for(int i=0;i<=endRow;i++){
			builder.append("\n  ").append(i).append("\t");
			if(getRow(i).isNull()){
				builder.append("-*");
				continue;
			}
			int endCell = getEndCellIndex(i);
			for(int j=0;j<=endCell;j++){
				SCell cell = getCell(i, j);
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
		builder.append("\n}\n");
	}

	@Override
	public void insertColumn(int columnIdx, int lastColumnIdx) {
		checkOrphan();
		if(columnIdx>lastColumnIdx){
			throw new IllegalArgumentException(columnIdx+">"+lastColumnIdx);
		}

		int size = lastColumnIdx - columnIdx + 1;
		insertAndSplitColumnArray(columnIdx,size);
		
		for(AbstractRowAdv row:rows.values()){
			row.insertCell(columnIdx,size);
		}
		Map<String,Object> dataBefore = shiftBeforeColumnInsert(columnIdx,lastColumnIdx);
		ModelUpdateUtil.addInsertDeleteUpdate(this, true, false, columnIdx, lastColumnIdx);
		shiftAfterColumnInsert(dataBefore,columnIdx,lastColumnIdx);
	}
	
	private Map<String, Object> shiftBeforeColumnInsert(int columnIdx,
			int lastColumnIdx) {
		Map<String,Object> dataBefore = new HashMap<String, Object>();
		// handling merged regions shift
		// find merge cells in affected region and remove them
		CellRegion affectedRegion = new CellRegion(0, columnIdx, book.getMaxRowIndex(), book.getMaxColumnIndex());
		List<CellRegion> toExtend = getOverlapsMergedRegions(affectedRegion, true);
		List<CellRegion> toShift = getContainsMergedRegions(affectedRegion);
		removeMergedRegion(affectedRegion, true);
		dataBefore.put("toExtend", toExtend);
		dataBefore.put("toShift", toShift);
		return dataBefore;
	}

	private void shiftAfterColumnInsert(Map<String, Object> dataBefore, int columnIdx, int lastColumnIdx) {
		int size = lastColumnIdx - columnIdx+1;
		// handling pic shift
		for (AbstractPictureAdv pic : pictures) {
			ViewAnchor anchor = pic.getAnchor();
			int idx = anchor.getColumnIndex();
			if (idx >= columnIdx) {
				anchor.setColumnIndex(idx + size);
			}
		}
		// handling chart shift
		for (AbstractChartAdv chart : charts) {
			ViewAnchor anchor = chart.getAnchor();
			int idx = anchor.getColumnIndex();
			if (idx >= columnIdx) {
				anchor.setColumnIndex(idx + size);
			}
		}
		
		// handling merged regions shift
		List<CellRegion> toExtend = (List<CellRegion>)dataBefore.get("toExtend");
		List<CellRegion> toShift = (List<CellRegion>)dataBefore.get("toShift");
		// extend/move removed merged cells, then add them back
		for(CellRegion r : toExtend) {
			addMergedRegion(new CellRegion(r.row, r.column, r.lastRow, r.lastColumn + size));
		}
		for(CellRegion r : toShift) {
			addMergedRegion(new CellRegion(r.row, r.column + size, r.lastRow, r.lastColumn + size));
		}

		//TODO shift data validation?
		
		extendFormula(new CellRegion(0,columnIdx,book.getMaxRowIndex(),lastColumnIdx), true);
		
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
		int size = lastColumnIdx - columnIdx + 1;
		deleteAndShrinkColumnArray(columnIdx,size);
		
		for(AbstractRowAdv row:rows.values()){
			row.deleteCell(columnIdx,size);
		}
		Map<String,Object> dataBefore = shiftBeforeColumnDelete(columnIdx,lastColumnIdx);
		ModelUpdateUtil.addInsertDeleteUpdate(this, false, false, columnIdx, lastColumnIdx);
		shiftAfterColumnDelete(dataBefore,columnIdx,lastColumnIdx);
	}
	
	private Map<String,Object> shiftBeforeColumnDelete(int columnIdx,int lastColumnIdx) {
		Map<String,Object> dataBefore = new HashMap<String,Object>();
		
		// handling merged regions shift
		// find merge cells in affected region and remove them
		CellRegion affectedRegion = new CellRegion(0, columnIdx, book.getMaxRowIndex(), book.getMaxColumnIndex());
		List<CellRegion> toShrink = getOverlapsMergedRegions(affectedRegion, false);
		removeMergedRegion(affectedRegion, true);
		dataBefore.put("toShrink", toShrink);
		
		return dataBefore;
	}
	private void shiftAfterColumnDelete(Map<String,Object> dataBefore, int columnIdx,int lastColumnIdx) {
		int size = lastColumnIdx - columnIdx+1;
		//handling pic shift
		for(AbstractPictureAdv pic:pictures){
			ViewAnchor anchor = pic.getAnchor();
			int idx = anchor.getColumnIndex();
			if(idx >= columnIdx+size){
				anchor.setColumnIndex(idx-size);
			}else if(idx >= columnIdx){
				anchor.setColumnIndex(columnIdx);//as excel's rule
				anchor.setXOffset(0);
			}
		}
		//handling chart shift
		for(AbstractChartAdv chart:charts){
			ViewAnchor anchor = chart.getAnchor();
			int idx = anchor.getColumnIndex();
			if(idx >= columnIdx+size){
				anchor.setColumnIndex(idx-size);
			}else if(idx >= columnIdx){
				anchor.setColumnIndex(columnIdx);//as excel's rule
				anchor.setXOffset(0);
			}
		}
		
		// handling merged regions shift
		List<CellRegion> toShrink = (List<CellRegion>)dataBefore.get("toShrink");
		// shrink/move removed merged cells, then add them back
		for(CellRegion r : toShrink) {
			CellRegion shrank = shrinkColumn(r, columnIdx, lastColumnIdx);
			if(shrank != null) {
				addMergedRegion(shrank);
			}
		}

		//TODO shift data validation?

		shrinkFormula(new CellRegion(0,columnIdx,book.getMaxRowIndex(),lastColumnIdx), true);
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
		
		//TODO zss 3.5 move to  movecellhelper
		//check merge overlaps and contains
		CellRegion srcRegion = new CellRegion(rowIdx,columnIdx,lastRowIdx,lastColumnIdx);
		Collection<CellRegion> containsMerge = getContainsMergedRegions(srcRegion);
		Collection<CellRegion> overlapsMerge = getOverlapsMergedRegions(srcRegion,true);
		if(overlapsMerge.size()>0){
			throw new InvalidateModelOpException("Can't move "+srcRegion.getReferenceString()+" which overlaps merge area "+overlapsMerge.iterator().next().getReferenceString());
		}
		CellRegion targetRegion = new CellRegion(rowIdx+rowOffset,columnIdx+columnOffset,lastRowIdx+rowOffset,lastColumnIdx+columnOffset);
		//to backward compatible with old spec, we should auto ummerge the target area
//		overlapsMerge = getOverlapsMergedRegions(targetRegion,true); 
//		if(overlapsMerge.size()>0){
//			throw new InvalidateModelOpException("Can't move to "+targetRegion.getReferenceString()+" which overlaps merge area");
//		}
		this.removeMergedRegion(targetRegion, true);
		
		
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
				SCell cell = row==null?null:row.getCell(c, false);
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
		
		//should use precedent update since the value might be changed and need to clear cache
		ModelUpdateUtil.handlePrecedentUpdate(getBook().getBookSeries(),
				new RefImpl(getBook().getBookName(), getSheetName(), rowIdx,
						columnIdx, lastRowIdx, lastColumnIdx));
		ModelUpdateUtil.handlePrecedentUpdate(getBook().getBookSeries(),
				new RefImpl(getBook().getBookName(), getSheetName(), rowIdx
						+ rowOffset, columnIdx + columnOffset, lastRowIdx
						+ rowOffset, lastColumnIdx + columnOffset));
		
		//shift the merge
		mergedRegions.removeAll(containsMerge);
		for(CellRegion merge:containsMerge){
			CellRegion newMerge = new CellRegion(merge.getRow() + rowOffset,merge.getColumn()+ columnOffset,
					merge.getLastRow()+rowOffset,merge.getLastColumn()+columnOffset);
			mergedRegions.add(newMerge);
			ModelUpdateUtil.addMergeUpdate(this,merge, newMerge);
		}
		
		shiftAfterCellMove(rowIdx, columnIdx,lastRowIdx,lastColumnIdx, rowOffset, columnOffset);
		
		//TODO validation and other stuff
		
		//TODO event
	}

	
	private void shiftAfterCellMove(int rowIdx, int columnIdx, int lastRowIdx,
			int lastColumnIdx, int rowOffset, int columnOffset) {
		moveFormula(new CellRegion(rowIdx,columnIdx,lastRowIdx,lastColumnIdx),rowOffset,columnOffset);
	}
	
	private void moveFormula(CellRegion src,int rowOffset,int columnOffset){
		SBook book = getBook();
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)book.getBookSeries();
		DependencyTable dt = bs.getDependencyTable();
		Set<Ref> dependents = dt.getDirectDependents(new RefImpl(book.getBookName(),getSheetName(),src.getRow(),src.getColumn(),src.getLastRow(),src.getLastColumn()));
		if(dependents.size()>0){
			FormulaTunerHelper tuner = new FormulaTunerHelper(bs);
			tuner.move(new SheetRegion(this,src),dependents,rowOffset,columnOffset);
		}
	}
	
	private void shrinkFormula(CellRegion src,boolean horizontal){
		SBook book = getBook();
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)book.getBookSeries();
		DependencyTable dt = bs.getDependencyTable();
		
		Ref ref = new RefImpl(book.getBookName(),getSheetName(),src.getRow(), src.getColumn(),
				horizontal?src.getLastRow():book.getMaxRowIndex(),horizontal?book.getMaxColumnIndex():src.getLastColumn());
		
		Set<Ref> dependents = dt.getDirectDependents(ref);
		if(dependents.size()>0){
			FormulaTunerHelper tuner = new FormulaTunerHelper(bs);
			tuner.shrink(new SheetRegion(this,src),dependents,horizontal);
		}
	}
	
	private void extendFormula(CellRegion src,boolean horizontal){
		SBook book = getBook();
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)book.getBookSeries();
		DependencyTable dt = bs.getDependencyTable();
		Ref ref = new RefImpl(book.getBookName(),getSheetName(),src.getRow(), src.getColumn(),
				horizontal?src.getLastRow():book.getMaxRowIndex(),horizontal?book.getMaxColumnIndex():src.getLastColumn());
		Set<Ref> dependents = dt.getDirectDependents(ref);
		if(dependents.size()>0){
			FormulaTunerHelper tuner = new FormulaTunerHelper(bs);
			tuner.extend(new SheetRegion(this,src),dependents,horizontal);
		}
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

	public SPicture addPicture(Format format, byte[] data,ViewAnchor anchor) {
		checkOrphan();
		AbstractPictureAdv pic = new PictureImpl(this,book.nextObjId("pic"), format, data,anchor);
		pictures.add(pic);
		return pic;
	}
	
	public SPicture getPicture(String picid){
		for(SPicture pic:pictures){
			if(pic.getId().equals(picid)){
				return pic;
			}
		}
		return null;
	}

	public void deletePicture(SPicture picture) {
		checkOrphan();
		checkOwnership(picture);
		((AbstractPictureAdv)picture).destroy();
		pictures.remove(picture);
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public List<SPicture> getPictures() {
		return Collections.unmodifiableList((List)pictures);
	}
	
	
	@Override
	public int getNumOfPicture() {
		return pictures.size();
	}

	@Override
	public SPicture getPicture(int idx) {
		return pictures.get(idx);
	}
	
	@Override
	public SChart addChart(SChart.ChartType type,ViewAnchor anchor) {
		checkOrphan();
		AbstractChartAdv pic = new ChartImpl(this, book.nextObjId("chart"), type, anchor);
		charts.add(pic);
		return pic;
	}
	@Override
	public SChart getChart(String picid){
		for(SChart pic:charts){
			if(pic.getId().equals(picid)){
				return pic;
			}
		}
		return null;
	}
	@Override
	public void deleteChart(SChart chart) {
		checkOrphan();
		checkOwnership(chart);
		((AbstractChartAdv)chart).destroy();
		charts.remove(chart);
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	@Override
	public List<SChart> getCharts() {
		return Collections.unmodifiableList((List)charts);
	}

	@Override
	public int getNumOfChart() {
		return charts.size();
	}

	@Override
	public SChart getChart(int idx) {
		return charts.get(idx);
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public List<CellRegion> getMergedRegions() {
		return Collections.unmodifiableList((List)mergedRegions);
	}

	@Override
	public void removeMergedRegion(CellRegion region,boolean removeOverlaps) {
		for(CellRegion r:new ArrayList<CellRegion>(mergedRegions)){
			if((removeOverlaps && region.overlaps(r)) || region.contains(r)){
				mergedRegions.remove(r);
				ModelUpdateUtil.addMergeUpdate(this,r, null);
			}
		}
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
		ModelUpdateUtil.addMergeUpdate(this,null, region);
	}

	@Override
	public List<CellRegion> getOverlapsMergedRegions(CellRegion region,boolean excludeContains){
		List<CellRegion> list =new LinkedList<CellRegion>(); 
		for(CellRegion r:mergedRegions){
			if(excludeContains && region.contains(r))
				continue;
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

	@SuppressWarnings("unchecked")
	@Override
	public Iterator<SRow> getRowIterator() {
		return Collections.unmodifiableCollection((Collection)rows.values()).iterator();
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public Iterator<SColumnArray> getColumnArrayIterator() {
		return Collections.unmodifiableCollection((Collection)columnArrays.values()).iterator();
	}
	

	@Override
	public Iterator<SColumn> getColumnIterator() {
		return new Iterator<SColumn>(){
			int index = -1;
			@Override
			public boolean hasNext() {
				return getColumnArray(index+1)!=null;
			}

			@Override
			public SColumn next() {
				index++;
				return getColumn(index);
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException("readonly");
			}
		};
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public Iterator<SCell> getCellIterator(int row) {
		return (Iterator)((AbstractRowAdv)getRow(row)).getCellIterator(false);
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
		return password!=null;
	}

	@Override
	public void setPassword(String password) {
		this.password = password;
	}
	
	@Override
	public String getPassword() {
		return this.password;
	}

	@Override
	public SSheetViewInfo getViewInfo(){
		return viewInfo;
	}
	
	@Override
	public SPrintSetup getPrintSetup(){
		return printSetup;
	}
	
	@Override
	public SDataValidation addDataValidation(CellRegion region) {
		return addDataValidation(region,null);
	}
	public SDataValidation addDataValidation(CellRegion region,SDataValidation src) {
		checkOrphan();
		Validations.argInstance(src, AbstractDataValidationAdv.class);
		AbstractDataValidationAdv validation = new DataValidationImpl(this, book.nextObjId("valid"));
		validation.setRegion(region);
		dataValidations.add(validation);
		if(src!=null){
			validation.copyFrom((AbstractDataValidationAdv)src);
		}
		return validation;
	}
	@Override
	public SDataValidation getDataValidation(String validationid){
		for(SDataValidation validation:dataValidations){
			if(validation.getId().equals(validationid)){
				return validation;
			}
		}
		return null;
	}
	@Override
	public void deleteDataValidation(SDataValidation validationid) {
		checkOrphan();
		checkOwnership(validationid);
		((AbstractDataValidationAdv)validationid).destroy();
		dataValidations.remove(validationid);
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public List<SDataValidation> getDataValidations() {
		return Collections.unmodifiableList((List)dataValidations);
	}

	@Override
	public int getNumOfDataValidation() {
		return dataValidations.size();
	}

	@Override
	public SDataValidation getDataValidation(int idx) {
		return dataValidations.get(idx);
	}
	
	@Override
	public SDataValidation getDataValidation(int rowIdx,int columnIdx) {
		for(SDataValidation validation:dataValidations){
			CellRegion region = validation.getRegion();
			if(region.contains(rowIdx, columnIdx)){
				return validation;
			}
		}
		return null;
	}

	@Override
	public SAutoFilter getAutoFilter() {
		return autoFilter;
	}

	@Override
	public SAutoFilter createAutoFilter(CellRegion region) {
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


	@Override
	public CellRegion pasteCell(SheetRegion src, CellRegion dest, PasteOption option) {
		return new PasteCellHelper(this).pasteCell(src,dest,option);
	}
}
