package org.zkoss.zss.ngapi.impl;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.ErrorConstants;
import org.zkoss.poi.xssf.usermodel.XSSFCell;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;

public abstract class AbstractImporter implements NImporter{

	@Override
	public NBook imports(File file, String bookName) throws IOException {
		InputStream is = null;
		try{
			is = new BufferedInputStream(new FileInputStream(file));
			return imports(is,bookName);
		}finally{
			if(is!=null){
				try{
					is.close();
				}catch(Exception x){};
			}
		}
	}

	@Override
	public NBook imports(URL url, String bookName) throws IOException {
		InputStream is = null;
		try {
			is = url.openStream();
			return imports(is, bookName);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (Exception x) {
				}
				;
			}
		}
	}
	
	protected NCell importPoiCell(NSheet sheet, Cell poiCell){
		NCell cell = sheet.getCell(poiCell.getRowIndex(), poiCell.getColumnIndex());
		switch (poiCell.getCellType()){
			case Cell.CELL_TYPE_NUMERIC:
				cell.setNumberValue(poiCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setStringValue(poiCell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cell.setBooleanValue(poiCell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				cell.setFormulaValue(poiCell.getCellFormula());
				break;
			case Cell.CELL_TYPE_ERROR:
				cell.setErrorValue(convertErrorCode(poiCell.getErrorCellValue()));
				break;
			case Cell.CELL_TYPE_BLANK:
				//do nothing because spreadsheet model auto creates blank cells
				break;
			default:
				//TODO log "ignore a cell with unknown.
		}
		return cell;
	}

	protected ErrorValue convertErrorCode(byte errorCellValue) {
		switch (errorCellValue){
			case ErrorConstants.ERROR_NAME:
				return new ErrorValue(ErrorValue.INVALID_NAME);
			case ErrorConstants.ERROR_VALUE:
				return new ErrorValue(ErrorValue.INVALID_VALUE);
			default:
				//TODO log it
				return new ErrorValue(ErrorValue.INVALID_NAME);
		}
		
	}
}
