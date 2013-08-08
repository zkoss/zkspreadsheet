/* CellDataImpl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api.model.impl;

import java.util.Date;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.impl.RangeImpl;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class CellDataImpl implements CellData{

	private RangeImpl _range;
	
	private Cell _cell;
	private boolean _cellInit;
	
	public CellDataImpl(RangeImpl range) {
		this._range = range;
	}

	@Override
	public int getRow() {
		return _range.getRow();
	}

	@Override
	public int getColumn() {
		return _range.getColumn();
	}

	private void initCell(){
		if(_cellInit){
			return;
		}
		_cellInit = true;
		XRange x = _range.getNative();
		XSheet sheet = x.getSheet();
		
		if(_cell==null){
			Row row = sheet.getRow(_range.getRow());
			if(row!=null){
				_cell = row.getCell(_range.getColumn());
			}
		}
	}
	

	@Override
	public CellType getResultType() {
		CellType type = getType();
		
		if(type != CellType.FORMULA){
			return type;
		}
		
		return toCellType(_cell.getCachedFormulaResultType());
	}
	
	private CellType toCellType(int type){
		switch(type){
		case Cell.CELL_TYPE_BLANK:
			return CellType.BLANK;
		case Cell.CELL_TYPE_BOOLEAN:
			return CellType.BOOLEAN;
		case Cell.CELL_TYPE_ERROR:
			return CellType.ERROR;
		case Cell.CELL_TYPE_FORMULA:
			return CellType.FORMULA;
		case Cell.CELL_TYPE_NUMERIC:
			return CellType.NUMERIC;
		case Cell.CELL_TYPE_STRING:
			return CellType.STRING;
		}
		return CellType.BLANK;
	}
	
	@Override
	public CellType getType() {
		initCell();
		
		if(_cell==null){
			return CellType.BLANK;
		}
		
		CellType type = toCellType(_cell.getCellType());
		
		if(type==CellType.FORMULA){
			if(toCellType(_cell.getCachedFormulaResultType())==CellType.ERROR){
				return CellType.ERROR;
			}
		}
		return type;
	}

	@Override
	public Object getValue() {
		return _range.getCellValue();
	}

	@Override
	public String getFormatText() {
		return _range.getCellFormatText();
	}

	@Override
	public String getEditText() {
		return _range.getCellEditText();
	}

	@Override
	public void setValue(Object value) {
		_range.setCellValue(value);
	}

	@Override
	public void setEditText(String editText) {
		_range.setCellEditText(editText);
	}

	public boolean validateEditText(String editText){
		return _range.getNative().validate(editText)==null;
	}

	@Override
	public boolean isBlank() {
		return getType() == CellType.BLANK;
	}

	@Override
	public boolean isFormula() {
		return getType() == CellType.FORMULA;
	}

	@Override
	public Double getDoubleValue() {
		Object val = getValue();
		if(val instanceof Double){
			return (Double)val;
		}
		return null;
	}

	@Override
	public Date getDateValue() {
		Double val = getDoubleValue();
		return val==null?null:BookHelper.numberToDate(_range.getBook().getPoiBook(), val);
	}

	@Override
	public String getStringValue() {
		Object val = getValue();
		if(val instanceof String){
			return (String)val;
		}
		return null;
	}

	@Override
	public Boolean getBooleanValue() {
		Object val = getValue();
		if(val instanceof Boolean){
			return (Boolean)val;
		}
		return null;
	}
}
