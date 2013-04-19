package org.zkoss.zss.api;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Worksheet;

public class NGetter {

	NRange range;

	public NGetter(NRange range) {
		this.range = range;
	}

	/**
	 * get the first cell style of this range
	 * 
	 * @return cell style if cell is exist, the check row style and column cell style if cell not found, if row and column style is not exist, then return default style of sheet
	 */
	public NCellStyle getCellStyle() {
		Worksheet sheet = range.getNative().getSheet();
		Book book = sheet.getBook();
		
		int r = range.getRow();
		int c = range.getColumn();
		CellStyle style = null;
		Row row = sheet.getRow(r);
		if (row != null){
			Cell cell = row.getCell(c);
			
			if (cell != null){//cell style
				style = cell.getCellStyle();
			}
			if(style==null && row.isFormatted()){//row sytle
				style = row.getRowStyle();
			}
		}
		if(style==null){//col style
			style = sheet.getColumnStyle(c);
		}
		if(style==null){//default
			style = book.getCellStyleAt((short) 0);
		}
		
		return new NCellStyle(sheet.getBook(), style);		
	}

}
