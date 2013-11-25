package org.zkoss.zss.ngmodel.impl;

import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.Notify;
import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;
import org.zkoss.zss.ngmodel.sys.input.InputEngine;
import org.zkoss.zss.ngmodel.sys.input.InputParseContext;
import org.zkoss.zss.ngmodel.sys.input.InputResult;
import org.zkoss.zss.ngmodel.util.Validations;


public class NRangeImpl implements NRange {
	
	private final List<EffectedRange> rangeRefs = new ArrayList<EffectedRange>(1);
	
	private int _column = Integer.MAX_VALUE;
	private int _row = Integer.MAX_VALUE;
	private int _lastColumn = Integer.MIN_VALUE;
	private int _lastRow = Integer.MIN_VALUE;
	
	private Locale _locale; 
	
	private NRangeImpl() {
		//TODO from zss context
		_locale = Locale.getDefault();
	}
	public NRangeImpl(NSheet sheet) {
		this();
		addRangeRef(sheet, 0, 0, sheet.getBook().getMaxRowSize(), sheet.getBook().getMaxColumnSize());
	}
	public NRangeImpl(NSheet sheet, int row, int col) {
		this();
		addRangeRef(sheet, row, col, row, col);
	}
	
	public NRangeImpl(NSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		this();
		addRangeRef(sheet, tRow, lCol, bRow, rCol);
	}
	
	public void setLocale(Locale locale){
		Validations.argNotNull(locale);
		_locale = locale;
	}
	
	public Locale getLocale(){
		return _locale;
	}

	private void addRangeRef(NSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		Validations.argNotNull(sheet);
		rangeRefs.add(new EffectedRange(sheet,tRow,lCol,bRow,rCol));
		
		_column = Math.min(_column,lCol);
		_row = Math.min(_row, tRow);
		_lastColumn = Math.max(_lastColumn,rCol);
		_lastRow = Math.max(_lastRow, bRow);
		
	}
	public NSheet getSheet() {
		return rangeRefs.get(0)._sheet;
	}
	public int getRow() {
		return _row;
	}
	public int getColumn() {
		return _column;
	}
	public int getLastRow() {
		return _lastRow;
	}
	public int getLastColumn() {
		return _lastColumn;
	}
	
	private class EffectedRange {
		private final NSheet _sheet;
		private final int _column;
		private final int _row;
		private final int _lastColumn;
		private final int _lastRow;
		
		public EffectedRange(NSheet sheet,int row, int column, int lastRow, int lastColumn){
			_sheet = sheet;
			_row = row;
			_column = column;
			_lastRow = lastRow;
			_lastColumn = lastColumn;	
		}
	}
	
	static interface CellVisitor{
		boolean visit(NCell cell);
	}
	
	private void travelCells(CellVisitor visitor){
		LinkedHashSet<Ref> dependentSet = new LinkedHashSet<Ref>();
		LinkedHashSet<Ref> updateSet = new LinkedHashSet<Ref>();
		
		NBook book = rangeRefs.get(0)._sheet.getBook();
		NBookSeries bookSeries = book.getBookSeries();
		DependencyTable dependencyTable = ((AbstractBookSeries)bookSeries).getDependencyTable();
		
		String bookName = book.getBookName();
		
		for(EffectedRange r:rangeRefs){
			String sheetName = r._sheet.getSheetName();
			for(int i=r._row;i<=r._lastRow;i++){
				for (int j = r._column; j <= r._lastColumn; j++) {
					NCell cell = r._sheet.getCell(i, j);
					boolean update = visitor.visit(cell);
					if(update){
						Ref ref = new RefImpl(bookName, sheetName, i, j); 
						updateSet.add(ref);
						
						Set<Ref> dependent = dependencyTable.getDependents(ref);
						updateSet.addAll(dependent);
						dependentSet.addAll(dependent);
					}
				}
			}
		}
		
		//clear formula cache
		for(Ref ref:dependentSet){
			if(ref.getType()==RefType.CELL){
				NBook dependentBook = bookSeries.getBook(ref.getBookName());
				NSheet dependentSheet = dependentBook.getSheetByName(ref.getSheetName());
				NCell cell = dependentSheet.getCell(ref.getRow(), ref.getColumn());
				cell.clearFormulaResultCache();
			}else{//another type?
				
			}
		}
		
		//notify changes
		for(Ref ref:updateSet){
			if(ref.getType()==RefType.CELL){
				NBook notifyBook = bookSeries.getBook(ref.getBookName());
				NSheet notifySheet = notifyBook.getSheetByName(ref.getSheetName());
				NCell cell = notifySheet.getCell(ref.getRow(), ref.getColumn());
				((AbstractBook)notifyBook).sendEvent(ModelEvents.ON_CELL_UPDATED, 
						ModelEvents.PARAM_CELL, cell);
			}else{//another type?
				
			}
		}
	}
	
	private boolean euqlas(Object obj1,Object obj2){
		if(obj1==obj2){
			return true;
		}
		if(obj1!=null){
			return obj1.equals(obj2);
		}
		return false;
	}
	
	public void setValue(final Object value) {
		travelCells(new CellVisitor() {
			public boolean visit(NCell cell) {
				Object cellval = cell.getValue();
				if(euqlas(cellval,value)){
					return false;
				}
				cell.setValue(value);
				return true;
			}
		});
	}
	public void clear() {
		travelCells(new CellVisitor() {
			public boolean visit(NCell cell) {
				if(cell.isNull() || cell.getType()==CellType.BLANK){
					return false;
				}
				cell.clearValue();
				return true;
			}
		});
	}
	
	public void setEditText(String editText) {
		final InputEngine ie = EngineFactory.getInstance().createInputEngine();
		final InputResult result = ie.parseInput(editText==null?"":editText, new InputParseContext(_locale));
		
		travelCells(new CellVisitor() {
			public boolean visit(NCell cell) {
				Object cellval = cell.getValue();
				if(euqlas(cellval,result.getValue())){
					return false;
				}
				switch(result.getType()){
				case BLANK:
					cell.clearValue();
					break;
				case DATE:
					cell.setDateValue((Date)result.getValue());
					break;
				case FORMULA:
					cell.setFormulaValue((String)result.getValue());
					break;
				case NUMBER:
					cell.setNumberValue((Number)result.getValue());
					break;
				case STRING:
					cell.setStringValue((String)result.getValue());
					break;
				case ERROR:
				default :
					cell.setValue(result.getValue());
				}
				return true;
			}
		});
	}
}
