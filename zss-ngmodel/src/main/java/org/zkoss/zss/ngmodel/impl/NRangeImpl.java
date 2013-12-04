package org.zkoss.zss.ngmodel.impl;

import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.ReadWriteTask;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;
import org.zkoss.zss.ngmodel.sys.input.InputEngine;
import org.zkoss.zss.ngmodel.sys.input.InputParseContext;
import org.zkoss.zss.ngmodel.sys.input.InputResult;
import org.zkoss.zss.ngmodel.util.Validations;

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
		boolean visit(NCell cell);
	}

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

		// clear formula cache
		for (Ref dependent : dependentSet) {
			if (dependent.getType() == RefType.CELL) {
				NBook dependentBook = bookSeries.getBook(dependent
						.getBookName());
				NSheet dependentSheet = dependentBook.getSheetByName(dependent
						.getSheetName());
				NCell cell = dependentSheet.getCell(dependent.getRow(),
						dependent.getColumn());
				cell.clearFormulaResultCache();
			} else {// another type?

			}
		}

		// notify changes
		for (Ref update : updateSet) {
			if (update.getType() == RefType.CELL) {
				NBook notifyBook = bookSeries.getBook(update.getBookName());
				NSheet notifySheet = notifyBook.getSheetByName(update
						.getSheetName());
				NCell cell = notifySheet.getCell(update.getRow(),
						update.getColumn());
				((BookAdv) notifyBook).sendEvent(
						ModelEvents.ON_CELL_UPDATED, ModelEvents.PARAM_CELL,
						cell);
			} else {// another type?

			}
		}
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
				if (euqlas(cellval, result.getValue())) {
					return false;
				}
				switch (result.getType()) {
				case BLANK:
					cell.clearValue();
					break;
				case DATE:
					cell.setDateValue((Date) result.getValue());
					break;
				case BOOLEAN:
					cell.setBooleanValue((Boolean) result.getValue());
					break;
				case FORMULA:
					cell.setFormulaValue((String) result.getValue());
					break;
				case NUMBER:
					cell.setNumberValue((Number) result.getValue());
					break;
				case STRING:
					cell.setStringValue((String) result.getValue());
					break;
				case ERROR:
				default:
					cell.setValue(result.getValue());
				}
				return true;
			}
		}).doInWriteLock(getLock());
	}
}
