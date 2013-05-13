package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.impl.RangeImpl;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.impl.XUtils;

public class CellDataImpl implements CellData{

	RangeImpl range;
	
	Cell cell;
	boolean cellInit;
	
	public CellDataImpl(RangeImpl range) {
		this.range = range;
	}

	@Override
	public int getRow() {
		return range.getRow();
	}

	@Override
	public int getColumn() {
		return range.getColumn();
	}

	private void initCell(){
		if(cellInit){
			return;
		}
		cellInit = true;
		XRange x = range.getNative();
		XSheet sheet = x.getSheet();
		
		if(cell==null){
			Row row = sheet.getRow(range.getRow());
			if(row!=null){
				cell = row.getCell(range.getColumn());
			}
		}
	}
	

	@Override
	public CellType getResultType() {
		CellType type = getType();
		
		if(type != CellType.FORMULA){
			return type;
		}
		
		return toCellType(cell.getCachedFormulaResultType());
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
		
		if(cell==null){
			return CellType.BLANK;
		}
		
		CellType type = toCellType(cell.getCellType());
		
		if(type==CellType.FORMULA){
			if(toCellType(cell.getCachedFormulaResultType())==CellType.ERROR){
				return CellType.ERROR;
			}
		}
		return type;
	}

	@Override
	public Object getValue() {
		return range.getCellValue();
	}

	@Override
	public String getFormatText() {
		//I don't create my way, use the same way from Spreadsheet implementation as possible
		return XUtils.getCellFormatText(range.getNative().getSheet(), getRow(), getColumn());
	}

	@Override
	public String getEditText() {
		return range.getCellEditText();
	}

	@Override
	public void setValue(Object value) {
		range.setCellValue(value);
	}

	@Override
	public void setEditText(String editText) {
		range.setCellEditText(editText);
	}

	public boolean validateEditText(String editText){
		return range.getNative().validate(editText)==null;
	}
}
