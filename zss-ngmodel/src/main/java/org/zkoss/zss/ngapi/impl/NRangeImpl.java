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
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.BookAdv;
import org.zkoss.zss.ngmodel.impl.BookSeriesAdv;
import org.zkoss.zss.ngmodel.impl.ModelInternalEvents;
import org.zkoss.zss.ngmodel.impl.RefImpl;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;
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

	private Locale _locale;

	private NRangeImpl() {
		// TODO from zss context
		_locale = Locale.getDefault();
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
	@Override
	public void setLocale(Locale locale) {
		Validations.argNotNull(locale);
		_locale = locale;
	}
	@Override
	public Locale getLocale() {
		return _locale;
	}

	private void addRangeRef(NSheet sheet, int tRow, int lCol, int bRow,
			int rCol) {
		Validations.argNotNull(sheet);
		rangeRefs.add(new EffectedRegion(sheet, tRow, lCol, bRow, rCol));

		_column = Math.min(_column, lCol);
		_row = Math.min(_row, tRow);
		_lastColumn = Math.max(_lastColumn, rCol);
		_lastRow = Math.max(_lastRow, bRow);

	}
	
	
	public ReadWriteLock getLock(){
		return getSheet().getBook().getBookSeries().getLock();
	}

	
	private class CellVisitorTask extends ReadWriteTask{
		private CellVisitor visitor;
		
		private CellVisitorTask(CellVisitor visitor){
			this.visitor = visitor;
		}

		@Override
		public Object invoke() {
			travelCells(visitor);
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
			travelFirstCell(visitor);
			return null;
		}
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

	static interface CellVisitor {
		/**
		 * @param cell
		 * @return true if this cell has been updated after visit
		 */
		boolean visit(NCell cell);
	}

	/**
	 * travels all the cells in this range
	 * @param visitor
	 */
	private void travelCells(CellVisitor visitor) {
		LinkedHashSet<Ref> dependentSet = new LinkedHashSet<Ref>();
		LinkedHashSet<Ref> updateSet = new LinkedHashSet<Ref>();

		NBook book = rangeRefs.get(0)._sheet.getBook();
		NBookSeries bookSeries = book.getBookSeries();
		DependencyTable dependencyTable = ((BookSeriesAdv) bookSeries)
				.getDependencyTable();

		String bookName = book.getBookName();

		for (EffectedRegion r : rangeRefs) {
			String sheetName = r._sheet.getSheetName();
			CellRegion region = r.region;
			for (int i = region.row; i <= region.lastRow; i++) {
				for (int j = region.column; j <= region.lastColumn; j++) {
					NCell cell = r._sheet.getCell(i, j);
					boolean update = visitor.visit(cell);
					if (update) {
						Ref ref = new RefImpl(bookName, sheetName, i, j);
						updateSet.add(ref);

						Set<Ref> dependent = dependencyTable.getDependents(ref);
						updateSet.addAll(dependent);
						dependentSet.addAll(dependent);
					}
				}
			}
		}

		handleRefUpdate(bookSeries,dependentSet,updateSet);
	}
	
	private void handleRefUpdate(NBookSeries bookSeries,LinkedHashSet<Ref> dependentSet,
			LinkedHashSet<Ref> updateSet) {
		// clear formula cache
		for (Ref dependent : dependentSet) {
			System.out.println(">>> Dependent "+dependent);
			//clear the dependent's formula cache since the precedent is changed.
			if (dependent.getType() == RefType.CELL) {
				NBook dependentBook = bookSeries.getBook(dependent
						.getBookName());
				NSheet dependentSheet = dependentBook.getSheetByName(dependent
						.getSheetName());
				NCell cell = dependentSheet.getCell(dependent.getRow(),
						dependent.getColumn());
				cell.clearFormulaResultCache();
			} else {// TODO another

			}
		}

		// notify changes
		for (Ref update : updateSet) {
			System.out.println(">>> Update "+update);
			RefType type = update.getType();
			if (type == RefType.CELL || type == RefType.AREA) {
				NBook notifyBook = bookSeries.getBook(update.getBookName());
				NSheet notifySheet = notifyBook.getSheetByName(update
						.getSheetName());
				((BookAdv) notifyBook).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_CELL_CONTENT_CHANGE,notifyBook,notifySheet,
						new CellRegion(update.getRow(),update.getColumn(),update.getLastRow(),update.getLastColumn())));
			} else {// TODO another

			}
		}
	}

	/**
	 * travels first cell in this range.
	 * @param visitor
	 */
	private void travelFirstCell(CellVisitor visitor) {
		LinkedHashSet<Ref> dependentSet = new LinkedHashSet<Ref>();
		LinkedHashSet<Ref> updateSet = new LinkedHashSet<Ref>();

		NBook book = rangeRefs.get(0)._sheet.getBook();
		NBookSeries bookSeries = book.getBookSeries();
		DependencyTable dependencyTable = ((BookSeriesAdv) bookSeries)
				.getDependencyTable();

		String bookName = book.getBookName();

		if(rangeRefs.size()<=0)
			return;
		
		EffectedRegion r = rangeRefs.get(0);
		String sheetName = r._sheet.getSheetName();
		CellRegion region = r.region;
			
		if(region.row <0 || region.column <0)
			return;
			
		NCell cell = r._sheet.getCell(region.row, region.column);
		boolean update = visitor.visit(cell);
		if (update) {
			Ref ref = new RefImpl(bookName, sheetName, region.row, region.column);
			updateSet.add(ref);//always update this cell

			//check if any dependent on this cell
			Set<Ref> dependent = dependencyTable.getDependents(ref);
			updateSet.addAll(dependent);
			dependentSet.addAll(dependent);
		}

		handleRefUpdate(bookSeries,dependentSet,updateSet);
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
	public void clear() {
		new CellVisitorTask(new CellVisitor() {
			public boolean visit(NCell cell) {
				if (cell.isNull() || cell.getType() == CellType.BLANK) {
					return false;
				}
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
						: editText, cell.getCellStyle().getDataFormat(), new InputParseContext(_locale));
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
						cell.setNumberValue((Number) resultVal);
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
					NStyles.setDataFormat(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex(), format);
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
}
