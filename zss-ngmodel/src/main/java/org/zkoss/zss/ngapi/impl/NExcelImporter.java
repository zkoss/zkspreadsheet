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

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Iterator;
import java.util.List;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellStyle;
import org.zkoss.poi.POIXMLDocument;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.poifs.filesystem.POIFSFileSystem;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.xssf.model.StylesTable;
import org.zkoss.poi.xssf.usermodel.XSSFCell;
import org.zkoss.poi.xssf.usermodel.XSSFCellStyle;
import org.zkoss.poi.xssf.usermodel.XSSFFont;
import org.zkoss.poi.xssf.usermodel.XSSFRow;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColor;
import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.BookAdv;
import org.zkoss.zss.ngmodel.impl.BookImpl;
import org.zkoss.zss.ngmodel.impl.CellAdv;
import org.zkoss.zss.ngmodel.impl.CellImpl;
import org.zkoss.zss.ngmodel.impl.CellStyleAdv;
import org.zkoss.zss.ngmodel.impl.CellStyleImpl;
import org.zkoss.zss.ngmodel.impl.ColorImpl;
import org.zkoss.zss.ngmodel.impl.FontImpl;
import org.zkoss.zss.ngmodel.impl.RowAdv;
import org.zkoss.zss.ngmodel.impl.RowImpl;
import org.zkoss.zss.ngmodel.impl.SheetAdv;
import org.zkoss.zss.ngmodel.impl.SheetImpl;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NExcelImporter extends AbstractImporter{


	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {

		Workbook workbook = null;

		if (POIFSFileSystem.hasPOIFSHeader(is)) {
			workbook = new HSSFWorkbook(is);
		}
		if (POIXMLDocument.hasOOXMLHeader(is)) {
			workbook = new XSSFWorkbook(is);
		}

		if (workbook == null) {
			throw new IllegalArgumentException("InputStream was neither an OLE2 stream, nor an OOXML stream");
		}

		NBook book = new BookImpl(bookName);
		
		if(workbook instanceof HSSFWorkbook) {
			
		}
		if(workbook instanceof XSSFWorkbook) {
			importXSSFBook(book, (XSSFWorkbook) workbook);
		}
		
		return book;
	}

	/**
	 * Copy XSSFBook attributes into NBook
	 * 
	 * @param nBook
	 * @param xssfBook
	 */
	private void importXSSFBook(NBook nBook, XSSFWorkbook xssfBook) {
		
		// FIXME: Should I handle style in the beginning?
		/**
		 * Handle book style table
		 */
		//StylesTable styleTable = xssfBook.getStylesSource();		
		/*
		 * Copy XSSFFont into NFont.
		 * Create Font into table for each Font.
		 */		
//		for(XSSFFont xssfFont : styleTable.getFonts()) {
//			
//			NFont nFont = new FontImpl();
//			
//			NColor nColor = new ColorImpl(xssfFont.getXSSFColor().getRgb());
//			// FIXME
//			// Should I create color?
//			// nBook.createColor();
//			
//			nFont.setColor(nColor);
//			nFont.setBoldweight(NFont.Boldweight.values()[xssfFont.getBoldweight()]);
//			nFont.setHeight(xssfFont.getFontHeight());
//			nFont.setItalic(xssfFont.getItalic());
//			nFont.setName(xssfFont.getFontName());
//			nFont.setStrikeout(xssfFont.getStrikeout());
//			nFont.setTypeOffset(NFont.TypeOffset.values()[xssfFont.getTypeOffset()]);
//			nFont.setUnderline(NFont.Underline.values()[xssfFont.getUnderline()]);
//			
//			bookAdv.createFont(nFont ,true);
//		}
		
		// Go through all sheet in book
		for(XSSFSheet xssfSheet : xssfBook) {
			NSheet nSheet = nBook.createSheet(xssfSheet.getSheetName());
			importXSSFSheet(nSheet, xssfSheet);
		}
	}

	/**
	 * Copy XSSFSheet attributes into NSheet
	 * @param nSheet
	 * @param xssfSheet
	 */
	private void importXSSFSheet(NSheet nSheet, XSSFSheet xssfSheet) {
		
		for(Row row : xssfSheet) { // Go through each row
			
			XSSFRow xssfRow = (XSSFRow) row;
			NRow nRow = nSheet.getRow(xssfRow.getRowNum());
			
			for(Cell cell : xssfRow) { // Go through each cell
				
				XSSFCell xssfCell = (XSSFCell) cell;
				NCell nCell = nSheet.getCell(xssfCell.getRowIndex(), xssfCell.getColumnIndex());
				
				//importXSSFCellStyle(nCell.getCellStyle(), xssfCell.getCellStyle());
				
				// TODO: copy hyper link
				// nCell.getHyperlink();
				
			}
			
		}
		
	}
	
	/**
	 * copy XSSFCellStyle attributes into nCellStyle
	 * @param nCellStyle
	 * @param xssfCellStyle
	 */
	private void importXSSFCellStyle(NCellStyle nCellStyle, XSSFCellStyle xssfCellStyle) {
		
		XSSFFont xssfFont = xssfCellStyle.getFont();
		NFont nFont = new FontImpl();
		
		NColor nColor = new ColorImpl(xssfFont.getXSSFColor().getRgb());
		
		// FIXME
		// Should I create color?
		// nBook.createColor();
		
		nFont.setColor(nColor);
		// FIXME: ENUM
		nFont.setBoldweight(NFont.Boldweight.values()[xssfFont.getBoldweight()]);
		nFont.setHeight(xssfFont.getFontHeight());
		nFont.setItalic(xssfFont.getItalic());
		nFont.setName(xssfFont.getFontName());
		nFont.setStrikeout(xssfFont.getStrikeout());
		// FIXME: ENUM
		nFont.setTypeOffset(NFont.TypeOffset.values()[xssfFont.getTypeOffset()]);
		// FIXME: ENUM
		nFont.setUnderline(NFont.Underline.values()[xssfFont.getUnderline()]);
		
		// FIXME
		// Should I create font?
		// nBook.createFont();
		
	}

}
