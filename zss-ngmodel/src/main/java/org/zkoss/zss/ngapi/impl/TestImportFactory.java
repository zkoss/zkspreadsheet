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

import org.zkoss.zss.ngapi.ImporterFactory;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.impl.BookImpl;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class TestImportFactory implements ImporterFactory{

	@Override
	public NImporter createImporter() {
		return new AbstractImporter() {
			
			@Override
			public NBook imports(InputStream is, String bookName) throws IOException {
				BookImpl book = new BookImpl(bookName);
				
				NSheet sheet1 = book.createSheet("Sheet 1");
				
				NCell cell = sheet1.getCell(3, 3);
				cell.setStringValue("ABCD");
				NCellStyle style = book.createCellStyle(true);
				style.setFont(book.createFont(true));
				style.getFont().setColor(book.createColor("#FF0000"));
				style.getFont().setHeight(16);
				style.setBackgroundColor(book.createColor("#0000AA"));
				
				
				sheet1.getColumn(3).setWidth(300);
				sheet1.getRow(3).setHeight(200);
				
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
				
				return book;
			}
		};
	}

}
