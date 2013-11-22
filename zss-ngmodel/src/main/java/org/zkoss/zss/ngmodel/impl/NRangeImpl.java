package org.zkoss.zss.ngmodel.impl;

import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.Notify;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.input.InputEngine;
import org.zkoss.zss.ngmodel.sys.input.InputParseContext;
import org.zkoss.zss.ngmodel.sys.input.InputResult;
import org.zkoss.zss.ngmodel.util.Validations;


public class NRangeImpl implements NRange {
//	private final NSheet _sheet;
	
	private List<RangeRef> rangeRefs = new ArrayList<RangeRef>(1);
	
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
		rangeRefs.add(new RangeRef(sheet,tRow,lCol,bRow,rCol));
		
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
	
	private InputEngine getInputEngine(){
		return EngineFactory.getInstance().getInputEngine();
	}

	public void setEditText(String editText) {
		if(editText==null)
			editText = "";
		
		InputEngine ie = getInputEngine();
		InputResult result = ie.parseInput(editText, new InputParseContext(_locale));
		
		LinkedHashSet<Notify> notifies = new LinkedHashSet<Notify>();
		for(RangeRef r:rangeRefs){
			for(int i=r._row;i<=r._lastRow;i++){
				for (int j = r._column; j <= r._lastColumn; j++) {
					NCell cell = r._sheet.getCell(i, j);
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
				}
			}
		}
		
		
	}
	
	private class RangeRef {
		private final NSheet _sheet;
		private final int _column;
		private final int _row;
		private final int _lastColumn;
		private final int _lastRow;
		
		public RangeRef(NSheet sheet,int row, int column, int lastRow, int lastColumn){
			_sheet = sheet;
			_row = row;
			_column = column;
			_lastRow = lastRow;
			_lastColumn = lastColumn;	
		}
	}

	public void setValue(Object value) {
		for(RangeRef r:rangeRefs){
			for(int i=r._row;i<=r._lastRow;i++){
				for (int j = r._column; j <= r._lastColumn; j++) {
					NCell cell = r._sheet.getCell(i, j);
					cell.setValue(value);
				}
			}
		}
	}
	public void clear() {
		for(RangeRef r:rangeRefs){
			for(int i=r._row;i<=r._lastRow;i++){
				for (int j = r._column; j <= r._lastColumn; j++) {
					NCell cell = r._sheet.getCell(i, j);
					cell.clearValue();
				}
			}
		}
	}
}
