package org.zkoss.zss.api.impl;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.Range.CellType;
import org.zkoss.zss.api.Range.CellValueHelper;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.fn.UtilFns;
import org.zkoss.zss.ui.impl.XUtils;

/*package*/ class CellValueHelperImpl implements CellValueHelper{

	RangeImpl range;
	
	Cell cell;
	boolean cellInit;
	
	public CellValueHelperImpl(RangeImpl range) {
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
	public CellType getType() {
		initCell();
		
		if(cell==null){
			return CellType.BLANK;
		}
		
		switch(cell.getCellType()){
		case Cell.CELL_TYPE_BLANK:
			return CellType.BLANK;
		case Cell.CELL_TYPE_BOOLEAN:
			return CellType.BOOLEAN;
		case Cell.CELL_TYPE_ERROR:
			return CellType.ERROR;
		case Cell.CELL_TYPE_FORMULA:
			if(getValue() instanceof Byte){
				return CellType.ERROR;
			}else{
				return CellType.FORMULA;
			}
		case Cell.CELL_TYPE_NUMERIC:
			return CellType.NUMERIC;
		case Cell.CELL_TYPE_STRING:
			return CellType.STRING;
		}
		return CellType.BLANK;
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
	
}
