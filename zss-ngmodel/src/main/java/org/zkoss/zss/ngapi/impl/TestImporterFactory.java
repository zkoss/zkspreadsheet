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

import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.zkoss.zss.ngapi.ImporterFactory;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBooks;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NCellStyle.FillPattern;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.impl.BookImpl;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class TestImporterFactory implements ImporterFactory{

	static NBook book;//test test share
	
	@Override
	public NImporter createImporter() {
		return new AbstractImporter() {
			
			@Override
			public NBook imports(InputStream is, String bookName) throws IOException {
//				if(book!=null){
//					return book;
//				}
//				book = NBooks.createBook(bookName);
//				book.setShareScope("application");
				
				NBook book = NBooks.createBook(bookName);
				
				NSheet sheet1 = book.createSheet("Sheet 1");
				sheet1.getColumn(0).setWidth(120);
				sheet1.getColumn(1).setWidth(120);
				sheet1.getColumn(2).setWidth(120);
				sheet1.getColumn(3).setWidth(120);
				sheet1.getColumn(4).setWidth(120);
				sheet1.getColumn(5).setWidth(120);
				sheet1.getColumn(6).setWidth(120);
				
				sheet1.getCell(0,11).setStringValue("Column M,O is hidden");
				sheet1.getColumn(12).setHidden(true);
				sheet1.getColumn(14).setHidden(true);
				
				sheet1.getCell(15,0).setStringValue("Row 17,19 is hidden");
				sheet1.getRow(16).setHidden(true);
				sheet1.getRow(18).setHidden(true);
				
				NCellStyle style;
				NCell cell;
				SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
				Date now = new Date();
				Date dayonly = null;
				try {
					dayonly = sdf.parse(sdf.format(now));
				} catch (ParseException e) {
					e.printStackTrace();
				}
				
				(cell = sheet1.getCell(0, 0)).setValue("Values:");
				(cell = sheet1.getCell(0, 1)).setValue(123.45);
				(cell = sheet1.getCell(0, 2)).setValue(now);
				(cell = sheet1.getCell(0, 3)).setValue(Boolean.TRUE);
				
				(cell = sheet1.getCell(1, 0)).setValue("Number Format:");
				(cell = sheet1.getCell(1, 1)).setValue(33);
				style = book.createCellStyle(true);
				style.setDataFormat("0.00");
				cell.setCellStyle(style);

				(cell = sheet1.getCell(1, 2)).setValue(44.55);
				style = book.createCellStyle(true);
				style.setDataFormat("$#,##0.0");
				cell.setCellStyle(style);
				
				
				(cell = sheet1.getCell(1, 3)).setValue(77.88);
				style = book.createCellStyle(true);
				style.setDataFormat("0.00;[Red]0.00");
				cell.setCellStyle(style);
				
				(cell = sheet1.getCell(1, 4)).setValue(-77.88);
				style = book.createCellStyle(true);
				style.setDataFormat("0.00;[Red]0.00");
				cell.setCellStyle(style);
				
				
				(cell = sheet1.getCell(2, 0)).setValue("Date Format:");
				(cell = sheet1.getCell(2, 1)).setValue(dayonly);
				style = book.createCellStyle(true);
				style.setDataFormat("yyyy/m/d");
				cell.setCellStyle(style);
				
				(cell = sheet1.getCell(2, 2)).setValue(dayonly);
				style = book.createCellStyle(true);
				style.setDataFormat("m/d/yyy");
				cell.setCellStyle(style);
				
				(cell = sheet1.getCell(2, 3)).setValue(now);
				style = book.createCellStyle(true);
				style.setDataFormat("m/d/yy h:mm;@");
				cell.setCellStyle(style);
				
				(cell = sheet1.getCell(2, 4)).setValue(now);
				style = book.createCellStyle(true);
				style.setDataFormat("h:mm AM/PM;@");
				cell.setCellStyle(style);
				
				//
				(cell = sheet1.getCell(3, 0)).setValue("Formula:");
				(cell = sheet1.getCell(3, 1)).setNumberValue(1);
				(cell = sheet1.getCell(3, 2)).setNumberValue(2);
				(cell = sheet1.getCell(3, 3)).setNumberValue(3);
				(cell = sheet1.getCell(3, 4)).setFormulaValue("SUM(B4:D4)");
				
				(cell = sheet1.getCell(4, 0)).setStringValue("this is a long long long long long string");
				
				(cell = sheet1.getCell(5, 0)).setStringValue("merege A6:C6");
				sheet1.addMergedRegion(new CellRegion(5,0,5,2));
				(cell = sheet1.getCell(5, 3)).setStringValue("merege D6:E7");
				sheet1.addMergedRegion(new CellRegion(5,3,6,4));
				
				
				cell = sheet1.getCell(9, 6);
				cell.setStringValue("G9");
				style = book.createCellStyle(true);
				style.setFont(book.createFont(true));
				style.getFont().setColor(book.createColor("#FF0000"));
				style.getFont().setHeightPoints(16);
				style.setFillPattern(FillPattern.SOLID_FOREGROUND);
				style.setFillColor(book.createColor("#AAAAAA"));
				
				
				sheet1.getColumn(6).setWidth(150);
				sheet1.getRow(9).setHeight(100);
				
				style.setAlignment(Alignment.RIGHT);
				style.setVerticalAlignment(VerticalAlignment.CENTER);
				
				style.setBorderTop(BorderType.THIN);
				style.setBorderBottom(BorderType.THIN);
				style.setBorderLeft(BorderType.THIN);
				style.setBorderRight(BorderType.THIN);
				style.setBorderTopColor(book.createColor("#FF0000"));
				style.setBorderBottomColor(book.createColor("#FFFF00"));
				style.setBorderLeftColor(book.createColor("#FF00FF"));
				style.setBorderRightColor(book.createColor("#00FFFF"));
				
				cell.setCellStyle(style);
				
				//row/column style
				
				style = book.createCellStyle(true);
				style.setFillPattern(FillPattern.SOLID_FOREGROUND);
				style.setFillColor(book.createColor("#FFAAAA"));
				sheet1.getRow(17).setCellStyle(style);
				sheet1.getCell(17, 0).setStringValue("row style");
				style = book.createCellStyle(true);
				style.setFillPattern(FillPattern.SOLID_FOREGROUND);
				style.setFillColor(book.createColor("#AAFFAA"));
				sheet1.getColumn(17).setCellStyle(style);
				sheet1.getColumn(17).setWidth(100);
				sheet1.getCell(0, 17).setStringValue("column style");
				
				return book;
			}
		};
	}

}
