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
package org.zkoss.zss.ngapi.impl;

import java.util.*;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.lang.Strings;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NAutoFilter.FilterOp;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NHyperlink.HyperlinkType;
import org.zkoss.zss.ngmodel.impl.*;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.format.*;
import org.zkoss.zss.ngmodel.sys.input.*;
import org.zkoss.zss.ngmodel.util.*;
/**
 * Only those methods that set cell data, cell style, row (column) style, width, height, and hidden consider 3-D references. 
 * Others don't, just perform on first cell.
 * @author dennis
 * @since 3.5.0
 */
public class NRangeImpl implements NRange {

	private NBook book;
	private final List<EffectedRegion> rangeRefs = new ArrayList<EffectedRegion>(
			1);

	private int _column = Integer.MAX_VALUE;
	private int _row = Integer.MAX_VALUE;
	private int _lastColumn = Integer.MIN_VALUE;
	private int _lastRow = Integer.MIN_VALUE;

	public NRangeImpl(NBook book) {
		this.book = book;
	}
	
	public NRangeImpl(NSheet sheet) {
		addRangeRef(sheet, 0, 0, sheet.getBook().getMaxRowSize(), sheet
				.getBook().getMaxColumnSize());
	}

	public NRangeImpl(NSheet sheet, int row, int col) {
		addRangeRef(sheet, row, col, row, col);
	}

	public NRangeImpl(NSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		addRangeRef(sheet, tRow, lCol, bRow, rCol);
	}

	private void addRangeRef(NSheet sheet, int tRow, int lCol, int bRow,
			int rCol) {
		Validations.argNotNull(sheet);
		//TODO to support multiple sheet
		rangeRefs.add(new EffectedRegion(sheet, tRow, lCol, bRow, rCol));

		_column = Math.min(_column, lCol);
		_row = Math.min(_row, tRow);
		_lastColumn = Math.max(_lastColumn, rCol);
		_lastRow = Math.max(_lastRow, bRow);

	}
	
	
	public ReadWriteLock getLock(){
		return getBookSeries().getLock();
	}

	
	private class CellVisitorTask extends ReadWriteTask{
		private CellVisitor visitor;
		private boolean stop = false;
		
		private CellVisitorTask(CellVisitor visitor){
			this.visitor = visitor;
		}

		@Override
		public Object invoke() {
			travelCells(visitor);
			return null;
		}
	}
	
	NBookSeries getBookSeries(){
		return getBook().getBookSeries();
	}
	
	NBook getBook(){
		if(book==null){
			book = getSheet().getBook();
		}
		return book;
	}
	
	@Override
	public NSheet getSheet() {
		if(rangeRefs.size()<=0){
			throw new IllegalStateException("can find any effected range");
		}
		return rangeRefs.get(0)._sheet;
	}
	@Override
	public int getRow() {
		return _row;
	}
	@Override
	public int getColumn() {
		return _column;
	}
	@Override
	public int getLastRow() {
		return _lastRow;
	}
	@Override
	public int getLastColumn() {
		return _lastColumn;
	}

	private class EffectedRegion {
		private final NSheet _sheet;
		private final CellRegion region;

		public EffectedRegion(NSheet sheet, int row, int column, int lastRow,
				int lastColumn) {
			_sheet = sheet;
			region = new CellRegion(row, column,lastRow,lastColumn);
		}
	}

	static abstract class CellVisitor {
		/**
		 * @param cell
		 * @return true if the cell has updated after visit
		 */
		abstract boolean visit(NCell cell);
		
		boolean isStopVisit(){
			return false;
		}
	}
	static abstract class OneCellVisitor extends CellVisitor{
		@Override
		boolean isStopVisit(){
			return true;
		}
	}
	

	/**
	 * travels all the cells in this range
	 * @param visitor
	 */
	private void travelCells(CellVisitor visitor) {
		LinkedHashSet<Ref> notifySet = new LinkedHashSet<Ref>();

		NBookSeries bookSeries = getBookSeries();

		DependentUpdateCollector dependentCtx = new DependentUpdateCollector();
		DependentUpdateCollector oldDependentCtx = DependentUpdateCollector.getCurrent();
		
		FormulaCacheCleaner oldClearer = FormulaCacheCleaner.setCurrent(new FormulaCacheCleaner(bookSeries));
		try{
			DependentUpdateCollector.setCurrent(dependentCtx);
			
			for (EffectedRegion r : rangeRefs) {
				String bookName = r._sheet.getBook().getBookName();
				String sheetName = r._sheet.getSheetName();
				CellRegion region = r.region;
				loop1:
				for (int i = region.row; i <= region.lastRow; i++) {
					for (int j = region.column; j <= region.lastColumn; j++) {
						NCell cell = r._sheet.getCell(i, j);
						boolean update = visitor.visit(cell);
						if(update){
							Ref ref = new RefImpl(bookName, sheetName, i, j);
							notifySet.add(ref);
						}
						if(visitor.isStopVisit()){
							break loop1;
						}
					}
				}
			}
		}finally{
			notifySet.addAll(dependentCtx.getDependents());
			
			DependentUpdateCollector.setCurrent(oldDependentCtx);
			FormulaCacheCleaner.setCurrent(oldClearer);
		}

		if(notifySet.size()>0){
			handleRefNotifyContentChange(bookSeries,notifySet);
		}
	}

	private void handleRefNotifyContentChange(NBookSeries bookSeries,HashSet<Ref> notifySet) {
		// notify changes
		new RefNotifyContentChangeHelper(bookSeries).notifyContentChange(notifySet);
	}

	private boolean euqlas(Object obj1, Object obj2) {
		if (obj1 == obj2) {
			return true;
		}
		if (obj1 != null) {
			return obj1.equals(obj2);
		}
		return false;
	}
	@Override
	public void setValue(final Object value) {
		new CellVisitorTask(new CellVisitor() {
			public boolean visit(NCell cell) {
				Object cellval = cell.getValue();
				if (!euqlas(cellval, value)) {
					cell.setValue(value);
					return true;
				}
				return false;
			}
		}).doInWriteLock(getLock());
	}
	
	@Override
	public void clearContents() {
		new CellVisitorTask(new CellVisitor() {
			public boolean visit(NCell cell) {
				if (!cell.isNull()) {
					cell.setHyperlink(null);
					cell.clearValue();
					return true;
				}
				return false;
			}
		}).doInWriteLock(getLock());
	}

	static class ResultWrap<T> {
		T obj;
		public ResultWrap(){}
		public ResultWrap(T obj){
			this.obj = obj;
		}
		public T get() {
			return obj;
		}
		public void set(T obj) {
			this.obj = obj;
		}
	}
	
	@Override
	public void setEditText(final String editText) {
		final InputEngine ie = EngineFactory.getInstance().createInputEngine();
		final ResultWrap<InputResult> input = new ResultWrap<InputResult>();
		new CellVisitorTask(new CellVisitor() {
			public boolean visit(NCell cell) {
				InputResult result;
				if((result = input.get())==null){
					result = ie.parseInput(editText == null ? ""
						: editText, cell.getCellStyle().getDataFormat(), new InputParseContext(Locales.getCurrent()));
					input.set(result);
				}
				
				Object cellval = cell.getValue();
				Object resultVal = result.getValue();
				String format = result.getFormat();
				if (euqlas(cellval, resultVal)) {
					return false;
				}
				
				switch (result.getType()) {
				case BLANK:
					cell.clearValue();
					break;
				case BOOLEAN:
					cell.setBooleanValue((Boolean) resultVal);
					break;
				case FORMULA:
					cell.setFormulaValue((String) resultVal);
					break;
				case NUMBER:
					if(resultVal instanceof Date){
						cell.setDateValue((Date)resultVal);
					}else{
						cell.setNumberValue((Double) resultVal);
					}
					break;
				case STRING:
					cell.setStringValue((String) resultVal);
					break;
				case ERROR:
				default:
					cell.setValue(resultVal);
				}
				
				String oldFormat = cell.getCellStyle().getDataFormat();
				if(format!=null && NCellStyle.FORMAT_GENERAL.equals(oldFormat)){
					//if there is a suggested format and old format is not general
					StyleUtil.setDataFormat(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex(), format);
				}
				return true;
			}
		}).doInWriteLock(getLock());
	}

	@Override
	public String getEditText() {
		final ResultWrap<String> r = new ResultWrap<String>();
		new CellVisitorTask(new OneCellVisitor() {
			@Override
			public boolean visit(NCell cell) {
				FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
				r.set(fe.getEditText(cell, new FormatContext(Locales.getCurrent())));		
				return false;
			}
		}).doInReadLock(getLock());
		return r.get();
	}

	@Override
	public void notifyChange() {
		NBookSeries bookSeries = getBookSeries();
		LinkedHashSet<Ref> notifySet = new LinkedHashSet<Ref>();
		for (EffectedRegion r : rangeRefs) {
			String bookName = r._sheet.getBook().getBookName();
			String sheetName = r._sheet.getSheetName();
			CellRegion region = r.region;
			Ref ref = new RefImpl(bookName, sheetName, region.row, region.column,region.lastRow,region.lastColumn);
			notifySet.add(ref);
		}
		handleRefNotifyContentChange(bookSeries,notifySet);
	}
	
	@Override
	public boolean isWholeSheet(){
		return isWholeRow()&&isWholeColumn();
	}

	@Override
	public boolean isWholeRow() {
		return _column<=0 && _lastColumn>=getBook().getMaxColumnSize();
	}

	@Override
	public NRange getRows() {
		return new NRangeImpl(getSheet(), _row, 0, _lastRow,getBook().getMaxColumnSize());
	}

	@Override
	public void setRowHeight(final int heightPx) {
		setRowHeight(heightPx,true);
	}
	public void setRowHeight(final int heightPx, final boolean custom) {
		new ReadWriteTask() {
			@Override
			public Object invoke() {
				setRowHeightInLock(heightPx,null,custom);
				return null;
			}
		}.doInWriteLock(getLock());
	}
	private void setRowHeightInLock(Integer heightPx,Boolean hidden, Boolean custom){
		LinkedHashSet<SheetRegion> notifySet = new LinkedHashSet<SheetRegion>();

		for (EffectedRegion r : rangeRefs) {
			int maxcol = r._sheet.getBook().getMaxColumnSize();
			CellRegion region = r.region;
			
			for (int i = region.row; i <= region.lastRow; i++) {
				NRow row = r._sheet.getRow(i);
				if(heightPx!=null){
					row.setHeight(heightPx);
				}
				if(hidden!=null){
					row.setHidden(hidden);
				}
				if(custom!=null){
					row.setCustomHeight(custom);
				}
				notifySet.add(new SheetRegion(r._sheet,i,0,i,maxcol));
			}
		}

		new NotifyChangeHelper(this).notifySizeChange(notifySet);
	}

	@Override
	public boolean isWholeColumn() {
		return _row<=0 && _lastRow>=getBook().getMaxRowSize();
	}

	@Override
	public NRange getColumns() {
		return new NRangeImpl(getSheet(), 0, _column, getBook().getMaxRowSize(), _lastColumn);
	}

	@Override
	public void setColumnWidth(final int widthPx) {
		setColumnWidth(widthPx,true);
	}
	public void setColumnWidth(final int widthPx,final boolean custom) {
		new ReadWriteTask() {
			@Override
			public Object invoke() {
				setColumnWidthInLock(widthPx,null,custom);
				return null;
			}
		}.doInWriteLock(getLock());
	}
	private void setColumnWidthInLock(Integer widthPx,Boolean hidden, Boolean custom){
		LinkedHashSet<SheetRegion> notifySet = new LinkedHashSet<SheetRegion>();

		for (EffectedRegion r : rangeRefs) {
			int maxrow = r._sheet.getBook().getMaxRowSize();
			CellRegion region = r.region;
			
			for (int i = region.column; i <= region.lastColumn; i++) {
				NColumn column = r._sheet.getColumn(i);
				if(widthPx!=null){
					column.setWidth(widthPx);
				}
				if(hidden!=null){
					column.setHidden(hidden);
				}
				if(custom!=null){
					column.setCustomWidth(true);
				}
				notifySet.add(new SheetRegion(r._sheet,0,i,maxrow,i));
			}
		}
		new NotifyChangeHelper(this).notifySizeChange(notifySet);
	}

	@Override
	public NHyperlink getHyperlink() {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public NRange copy(NRange dstRange, boolean cut) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public NRange copy(NRange dstRange) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public NRange pasteSpecial(NRange dstRange, PasteType pasteType,
			PasteOperation pasteOp, boolean skipBlanks, boolean transpose) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void insert(InsertShift shift, InsertCopyOrigin copyOrigin) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void delete(DeleteShift shift) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void merge(boolean across) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void unmerge() {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void setBorders(ApplyBorderType borderIndex, BorderType lineStyle,
			String color) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void move(int nRow, int nCol) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public NRange getCells(int row, int col) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void setStyle(NCellStyle style) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void autoFill(NRange dstRange, AutoFillType fillType) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void fillDown() {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void fillLeft() {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void fillRight() {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void fillUp() {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void setHidden(final boolean hidden) {
		new ReadWriteTask() {
			@Override
			public Object invoke() {
				setHiddenInLock(hidden);
				return null;
			}
		}.doInWriteLock(getLock());
	}
	
	
	private boolean isWholeRow(NBook book,CellRegion region){
		return region.column<=0 && region.lastColumn>=book.getMaxColumnSize();
	}
	
	private boolean isWholeColumn(NBook book,CellRegion region){
		return region.row<=0 && region.lastRow>=book.getMaxRowSize();
	}

	protected void setHiddenInLock(boolean hidden) {
		LinkedHashSet<SheetRegion> notifySet = new LinkedHashSet<SheetRegion>();
		for (EffectedRegion r : rangeRefs) {
			NBook book = r._sheet.getBook();
			int maxcol = r._sheet.getBook().getMaxColumnSize();
			int maxrow = r._sheet.getBook().getMaxRowSize();
			CellRegion region = r.region;
			
			if(isWholeRow(book,region)){//hidden the row when it is whole row
				for(int i = region.getRow(); i<=region.getLastRow();i++){
					NRow row = r._sheet.getRow(i);
					if(row.isHidden()==hidden)
						continue;
					row.setHidden(hidden);
					notifySet.add(new SheetRegion(r._sheet,i,0,i,maxcol));
				}
			}else if(isWholeColumn(book,region)){
				for(int i = region.getColumn(); i<=region.getLastColumn();i++){
					NColumn col = r._sheet.getColumn(i);
					if(col.isHidden()==hidden)
						continue;
					col.setHidden(hidden);
					notifySet.add(new SheetRegion(r._sheet,0,i,maxrow,i));
				}
			}
		}
		new NotifyChangeHelper(this).notifySizeChange(notifySet);
	}

	@Override
	public void setDisplayGridlines(boolean show) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void protectSheet(String password) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void setHyperlink(HyperlinkType linkType, String address,
			String display) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public Object getValue() {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public NRange getOffset(int rowOffset, int colOffset) {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public boolean isAnyCellProtected() {
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void deleteSheet() {
		final ResultWrap<NSheet> toDeleteSheet = new ResultWrap<NSheet>();
		final ResultWrap<Integer> toDeleteIndex = new ResultWrap<Integer>();
		//it just handle the first ref
		new DependentUpdateTask() {			
			@Override
			public Object doInvokePhase() {
				NBook book = getBook();
				int sheetCount;
				if((sheetCount = book.getNumOfSheet())<=1){
					throw new InvalidateModelOpException("can't delete last sheet ");
				}
				
				NSheet toDelete = getSheet();
				
				int index = book.getSheetIndex(toDelete);
//				final int newIndex =  index < (sheetCount - 1) ? index : (index - 1);
				
				toDeleteSheet.set(toDelete);
				toDeleteIndex.set(index);
				
				book.deleteSheet(toDelete);
				return null;
			}

			@Override
			void doNotifyPhase() {
				if(toDeleteSheet.get()!=null){
					((AbstractBookAdv) getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_SHEET_DELETE,book,
						ModelEvents.createDataMap(ModelEvents.PARAM_SHEET,toDeleteSheet.get(),ModelEvents.PARAM_INDEX,toDeleteIndex.get())));
				}
			}
		}.doInWriteLock(getLock());
	}

	@Override
	public NSheet createSheet(final String name) {
		final ResultWrap<NSheet> resultSheet = new ResultWrap<NSheet>();
		//it just handle the first ref
		return (NSheet)new DependentUpdateTask() {			
			@Override
			public Object doInvokePhase() {
				NBook book = getBook();
				NSheet sheet;
				if (Strings.isBlank(name)) {
					sheet = book.createSheet(nextSheetName());
				} else {
					sheet = book.createSheet(name);
				}
				resultSheet.set(sheet);
				return sheet;
			}

			@Override
			void doNotifyPhase() {
				if(resultSheet.get()!=null){
					((AbstractBookAdv) getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_SHEET_CREATE,resultSheet.get()));
				}
			}
		}.doInWriteLock(getLock());	
	}
	
	private String nextSheetName() {
		NBook book = getBook();
		Integer idx = (Integer)book.getAttribute("zss.nextSheetCount");
		int i = idx==null?1:idx;
		HashSet<String> names = new HashSet<String>();
		for (NSheet sheet : getBook().getSheets()) {
			names.add(sheet.getSheetName());
		}
		String base = "Sheet";
		String name = base + i; 
		while (names.contains(name)) {
			name = base+ (++i);
		}
		book.setAttribute("zss.nextSheetCount", Integer.valueOf(i+1));
		return name;
	}

	@Override
	public void setSheetName(final String newname) {
		//it just handle the first ref
		final ResultWrap<NSheet> resultSheet = new ResultWrap<NSheet>();
		new DependentUpdateTask() {			
			@Override
			public Object doInvokePhase() {
				NBook book = getBook();
				NSheet sheet = getSheet();
				if(sheet.getSheetName().equals(newname)){
					return null;
				}
				book.setSheetName(sheet, newname);
				resultSheet.set(sheet);
				
				return null;
			}

			@Override
			void doNotifyPhase() {
				if(resultSheet.get()!=null){
					((AbstractBookAdv) getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_SHEET_NAME_CHANGE,resultSheet.get()));
				}
			}
		}.doInWriteLock(getLock());	
	}

	@Override
	public void setSheetOrder(final int pos) {
		//it just handle the first ref
		final ResultWrap<NSheet> resultSheet = new ResultWrap<NSheet>();
		new DependentUpdateTask() {			
			@Override
			public Object doInvokePhase() {
				NBook book = getBook();
				NSheet sheet = getSheet();
				
				int oldIdx = book.getSheetIndex(sheet);
				if(oldIdx==pos){
					return null;
				}
				
				//in our new model, we don't use sheet index, so we don't need to clear anything when move it
				book.moveSheetTo(sheet, pos);
				resultSheet.set(sheet);
				return null;
			}

			@Override
			void doNotifyPhase() {
				if(resultSheet.get()!=null){
					((AbstractBookAdv) getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_SHEET_ORDER_CHANGE,resultSheet.get()));
				}
			}
		}.doInWriteLock(getLock());	
	}

	@Override
	public void setFreezePanel(final int numOfRow, final int numOfColumn) {
		//first ref only
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				NSheetViewInfo viewInfo = getSheet().getViewInfo();
				viewInfo.setNumOfRowFreeze(numOfRow);
				viewInfo.setNumOfColumnFreeze(numOfColumn);
				notifySheetFreezeChange();
				return null;
			}
		}.doInWriteLock(getLock());	
	}
	
	private void notifySheetFreezeChange(){
		new NotifyChangeHelper(this).notifySheetFreezeChange(new SheetRegion(getSheet(),-1,-1,-1,-1));
	}

	@Override
	public String getCellFormatText() {
		final ResultWrap<String> r = new ResultWrap<String>();
		new CellVisitorTask(new OneCellVisitor() {
			@Override
			public boolean visit(NCell cell) {
				FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
				r.set(fe.format(cell, new FormatContext(Locales.getCurrent())).getText());		
				return false;
			}
		}).doInReadLock(getLock());
		return r.get();
				
	}

	@Override
	public boolean isSheetProtected() {
		//TODO do we really need to use lock in such simple call, it looks overkill.
		return (Boolean)new ReadWriteTask(){
			@Override
			public Object invoke() {
				return getSheet().isProtected();
			}}.doInReadLock(getLock());
	}

	@Override
	public NDataValidation validate(final String editText) {
		final ResultWrap<NDataValidation> retrunVal = new ResultWrap<NDataValidation>();
		new CellVisitorTask(new CellVisitor() {
			boolean visit(NCell cell) {
				NDataValidation validation = getSheet().getDataValidation(cell.getRowIndex(), cell.getColumnIndex());
				if(validation!=null){
					if(!new DataValidationHelper(validation).validate(editText,cell.getCellStyle().getDataFormat())){
						retrunVal.set(validation);
					}
				}
				return false;
			}
			
			@Override
			boolean isStopVisit(){
				return retrunVal.get()!=null;
			}
			
		}).doInReadLock(getLock());
		return retrunVal.get();
	}

	@Override
	public NRange findAutoFilterRange() {
		return (NRange) new ReadWriteTask() {			
			@Override
			public Object invoke() {
				CellRegion region = new DataRegionHelper(NRangeImpl.this).findAutoFilterDataRegion();
				if(region!=null){
					return NRanges.range(getSheet(),region.getRow(),region.getColumn(),region.getLastRow(),region.getLastColumn());
				}
				return null;
			}
		}.doInReadLock(getLock());
	}
	
	Ref getSheetRef(){
		return new RefImpl((AbstractSheetAdv)getSheet());
	}
	Ref getBookRef(){
		return new RefImpl((AbstractSheetAdv)getBook());
	}
	
	private Set<Ref> toSet(Ref ref){
		Set<Ref> refs = new HashSet(1);
		refs.add(ref);
		return refs;
	}
	
	@Override 
	public NAutoFilter enableAutoFilter(final boolean enable){
		//it just handle the first ref
		return (NAutoFilter) new ReadWriteTask() {			
			@Override
			public Object invoke() {
				NSheet sheet = getSheet();
				NAutoFilter filter = sheet.getAutoFilter();
				
				if((filter==null && !enable) || (filter!=null && enable)){
					return filter;
				}
				
				filter = new AutoFilterHelper(NRangeImpl.this).enableAutoFilter(enable);
				notifySheetAutoFilterChange();
				return filter;
			}
		}.doInWriteLock(getLock());
	}

	@Override
	public NAutoFilter enableAutoFilter(final int field, final FilterOp filterOp,
			final Object criteria1, final Object criteria2, final Boolean visibleDropDown) {
		//it just handle the first ref
		return (NAutoFilter) new ReadWriteTask() {			
			@Override
			public Object invoke() {
				NAutoFilter filter = new AutoFilterHelper(NRangeImpl.this).enableAutoFilter(field, filterOp, criteria1, criteria2, visibleDropDown);
				notifySheetAutoFilterChange();
				return filter;
			}
		}.doInWriteLock(getLock());
	}
	
	public void resetAutoFilter(){
		//it just handle the first ref
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				new AutoFilterHelper(NRangeImpl.this).resetAutoFilter();
				notifySheetAutoFilterChange();
				return null;
			}
		}.doInWriteLock(getLock());		
	}
	
	public void applyAutoFilter(){
		//it just handle the first ref
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				new AutoFilterHelper(NRangeImpl.this).applyAutoFilter();
				notifySheetAutoFilterChange();
				return null;
			}
		}.doInWriteLock(getLock());		
	}
	
	private void notifySheetAutoFilterChange(){
		new NotifyChangeHelper(this).notifySheetAutoFilterChange(new SheetRegion(getSheet(),-1,-1,-1,-1));
	}

	@Override
	public void notifyCustomEvent(final String customEventName, final Object data, boolean writelock) {
		//it just handle the first ref
		ReadWriteTask task = new ReadWriteTask() {			
			@Override
			public Object invoke() {
				for (EffectedRegion r : rangeRefs) {
					NBook book = r._sheet.getBook();
					((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(customEventName,r._sheet,
							ModelEvents.createDataMap(ModelEvents.PARAM_CUSTOM_DATA,data)));
				}
				return null;
			}
		};
		if(writelock){
			task.doInWriteLock(getLock());
		}else{
			task.doInReadLock(getLock());
		}
	}
	
	private abstract class DependentUpdateTask extends ReadWriteTask{

		
		abstract Object doInvokePhase();
		abstract void doNotifyPhase();
		
		@Override
		public Object invoke() {
			LinkedHashSet<Ref> notifySet = new LinkedHashSet<Ref>();

			NBookSeries bookSeries = getBookSeries();

			DependentUpdateCollector dependentCtx = new DependentUpdateCollector();
			DependentUpdateCollector oldDependentCtx = DependentUpdateCollector.getCurrent();
			
			FormulaCacheCleaner oldClearer = FormulaCacheCleaner.setCurrent(new FormulaCacheCleaner(bookSeries));
			Object result = null;
			try{
				DependentUpdateCollector.setCurrent(dependentCtx);
				
				result = doInvokePhase();
			}finally{
				notifySet.addAll(dependentCtx.getDependents());
				
				DependentUpdateCollector.setCurrent(oldDependentCtx);
				FormulaCacheCleaner.setCurrent(oldClearer);
			}

			if(notifySet.size()>0){
				handleRefNotifyContentChange(bookSeries,notifySet);
			}
			doNotifyPhase();
			return result;
		}
	}

	@Override
	public NPicture addPicture(final NViewAnchor anchor, final byte[] image, final NPicture.Format format){
		return (NPicture) new ReadWriteTask() {			
			@Override
			public Object invoke() {
				NPicture picture = getSheet().addPicture(format, image, anchor);
				new NotifyChangeHelper(NRangeImpl.this).notifySheetPictureAdd(getSheet(), picture);
				return picture;
			}
		}.doInWriteLock(getLock());
	}
	
	@Override
	public void deletePicture(final NPicture picture){
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				getSheet().deletePicture(picture);
				new NotifyChangeHelper(NRangeImpl.this).notifySheetPictureDelete(getSheet(), picture);
				return null;
			}
		}.doInWriteLock(getLock());
	}
	
	@Override
	public void movePicture(final NPicture picture, final NViewAnchor anchor){
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				picture.setAnchor(anchor);
				new NotifyChangeHelper(NRangeImpl.this).notifySheetPictureMove(getSheet(), picture);
				return null;
			}
		}.doInWriteLock(getLock());
	}
}
