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

import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.poi.ss.formula.eval.NotImplementedException;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NDataValidation;
import org.zkoss.zss.ngmodel.NHyperlink;
import org.zkoss.zss.ngmodel.NHyperlink.HyperlinkType;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.format.FormatContext;
import org.zkoss.zss.ngmodel.sys.format.FormatEngine;
import org.zkoss.zss.ngmodel.sys.input.InputEngine;
import org.zkoss.zss.ngmodel.sys.input.InputParseContext;
import org.zkoss.zss.ngmodel.sys.input.InputResult;
import org.zkoss.zss.ngmodel.util.ReadWriteTask;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NRangeImpl implements NRange {

	private final List<EffectedRegion> rangeRefs = new ArrayList<EffectedRegion>(
			1);

	private int _column = Integer.MAX_VALUE;
	private int _row = Integer.MAX_VALUE;
	private int _lastColumn = Integer.MIN_VALUE;
	private int _lastRow = Integer.MIN_VALUE;

	private NRangeImpl() {
	}

	public NRangeImpl(NSheet sheet) {
		this();
		addRangeRef(sheet, 0, 0, sheet.getBook().getMaxRowSize(), sheet
				.getBook().getMaxColumnSize());
	}

	public NRangeImpl(NSheet sheet, int row, int col) {
		this();
		addRangeRef(sheet, row, col, row, col);
	}

	public NRangeImpl(NSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		this();
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
			travelCells(visitor,false);
			return null;
		}
	}
	
	private class FirstCellVisitorTask extends ReadWriteTask{
		private CellVisitor visitor;
		
		private FirstCellVisitorTask(CellVisitor visitor){
			this.visitor = visitor;
		}

		@Override
		public Object invoke() {
			travelCells(visitor,true);
			return null;
		}
	}
	
	NBookSeries getBookSeries(){
		return getBook().getBookSeries();
	}
	
	NBook getBook(){
		return getSheet().getBook();
	}
	
	@Override
	public NSheet getSheet() {
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
		 * @return true if this cell has been updated after visit
		 */
		abstract boolean visit(NCell cell);

		boolean isStopVisit(){
			return false;
		}
	}
	

	/**
	 * travels all the cells in this range
	 * @param visitor
	 */
	private void travelCells(CellVisitor visitor,boolean firstOnly) {
		LinkedHashSet<Ref> dependentSet = new LinkedHashSet<Ref>();
		LinkedHashSet<Ref> notifySet = new LinkedHashSet<Ref>();

		NBookSeries bookSeries = getBookSeries();
		DependencyTable dependencyTable = ((AbstractBookSeriesAdv) bookSeries)
				.getDependencyTable();

		for (EffectedRegion r : rangeRefs) {
			String bookName = r._sheet.getBook().getBookName();
			String sheetName = r._sheet.getSheetName();
			CellRegion region = r.region;
			loop1:
			for (int i = region.row; i <= region.lastRow; i++) {
				for (int j = region.column; j <= region.lastColumn; j++) {
					NCell cell = r._sheet.getCell(i, j);
					boolean update = visitor.visit(cell);
					if (update) {
						Ref ref = new RefImpl(bookName, sheetName, i, j);
						notifySet.add(ref);

						Set<Ref> dependent = dependencyTable.getDependents(ref);
						notifySet.addAll(dependent);
						dependentSet.addAll(dependent);
					}
					if(firstOnly){
						break loop1;
					}
					if(visitor.isStopVisit()){
						break loop1;
					}
				}
			}
		}

		handleRefDependent(bookSeries,dependentSet);
		handleRefNotifyContentChange(bookSeries,notifySet);
	}
	
	private void handleRefDependent(NBookSeries bookSeries,HashSet<Ref> dependentSet) {
		// clear formula cache
		new RefDependentHelper(bookSeries).handle(dependentSet);
	}
	
	private void handleRefNotifyContentChange(NBookSeries bookSeries,HashSet<Ref> notifySet) {
		// notify changes
		new RefNotifyContentChangeHelper(bookSeries).handle(notifySet);
	}
	
	private void handleRefNotifySizeChange(NBookSeries bookSeries,HashSet<Ref> notifySet) {
		new RefNotifySizeChangeHelper(bookSeries).handle(notifySet);
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
				if (euqlas(cellval, value)) {
					return false;
				}
				cell.setValue(value);
				return true;
			}
		}).doInWriteLock(getLock());
	}
	
	@Override
	public void clearContents() {
		new CellVisitorTask(new CellVisitor() {
			public boolean visit(NCell cell) {
				if (cell.isNull()) {
					return false;
				}
				cell.setHyperlink(null);
				cell.clearValue();
				return true;
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
		new FirstCellVisitorTask(new CellVisitor() {
			@Override
			public boolean visit(NCell cell) {
				FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
				r.set(fe.getEditText(cell, new FormatContext(Locales.getCurrent())));		
				return false;//no update
			}
		}).doInReadLock(getLock());
		return r.get();
	}

	@Override
	public void notifyChange() {
		NBook book = rangeRefs.get(0)._sheet.getBook();
		NBookSeries bookSeries = book.getBookSeries();
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
		LinkedHashSet<Ref> notifySet = new LinkedHashSet<Ref>();
		NBookSeries bookSeries = getBookSeries();

		for (EffectedRegion r : rangeRefs) {
			String bookName = r._sheet.getBook().getBookName();
			int maxcol = r._sheet.getBook().getMaxColumnSize();
			String sheetName = r._sheet.getSheetName();
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
				notifySet.add(new RefImpl(bookName,sheetName,i,0,i,maxcol));
			}
		}

		handleRefNotifySizeChange(bookSeries, notifySet);
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
		LinkedHashSet<Ref> notifySet = new LinkedHashSet<Ref>();
		NBookSeries bookSeries = getBookSeries();

		for (EffectedRegion r : rangeRefs) {
			String bookName = r._sheet.getBook().getBookName();
			int maxrow = r._sheet.getBook().getMaxRowSize();
			String sheetName = r._sheet.getSheetName();
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
				notifySet.add(new RefImpl(bookName,sheetName,0,i,maxrow,i));
			}
		}
		handleRefNotifySizeChange(bookSeries, notifySet);
	}

	@Override
	public NHyperlink getHyperlink() {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public NRange copy(NRange dstRange, boolean cut) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public NRange copy(NRange dstRange) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public NRange pasteSpecial(NRange dstRange, PasteType pasteType,
			PasteOperation pasteOp, boolean skipBlanks, boolean transpose) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void insert(InsertShift shift, InsertCopyOrigin copyOrigin) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void delete(DeleteShift shift) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void merge(boolean across) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void unmerge() {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void setBorders(ApplyBorderType borderIndex, BorderType lineStyle,
			String color) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void move(int nRow, int nCol) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public NRange getCells(int row, int col) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void setStyle(NCellStyle style) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void autoFill(NRange dstRange, AutoFillType fillType) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void fillDown() {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void fillLeft() {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void fillRight() {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void fillUp() {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void setHidden(boolean hidden) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void setDisplayGridlines(boolean show) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void protectSheet(String password) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void setHyperlink(HyperlinkType linkType, String address,
			String display) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public Object getValue() {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public NRange getOffset(int rowOffset, int colOffset) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public boolean isAnyCellProtected() {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void deleteSheet() {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public NSheet createSheet(String name) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void setSheetName(String name) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void setSheetOrder(int pos) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public void setFreezePanel(int rowfreeze, int columnfreeze) {
		throw new NotImplementedException("not implement yet");
	}

	@Override
	public String getCellFormatText() {
		final ResultWrap<String> r = new ResultWrap<String>();
		new FirstCellVisitorTask(new CellVisitor() {
			@Override
			public boolean visit(NCell cell) {
				FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
				r.set(fe.format(cell, new FormatContext(Locales.getCurrent())).getText());		
				return false;//no update
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
			public boolean visit(NCell cell) {
				NDataValidation validation = getSheet().getDataValidation(cell.getRowIndex(), cell.getColumnIndex());
				if(validation!=null){
					if(!new DataValidationHelper(validation).validate(editText,cell.getCellStyle().getDataFormat())){
						retrunVal.set(validation);
					}
				}
				return false;
			}

			@Override
			boolean isStopVisit() {
				//break visit if there is any validation
				return retrunVal.get()!=null;
			}
			
		}).doInReadLock(getLock());
		return retrunVal.get();
	}
}
