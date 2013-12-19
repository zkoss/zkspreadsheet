/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by Hawk
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl.imexp;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.ss.usermodel.Name;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NName;
import org.zkoss.zss.ngmodel.NSheet;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NExcelXlsExporter extends AbstractExcelExporter {
	
	@Override
	public void export(NBook book, OutputStream fos) throws IOException {
		ReadWriteLock lock = book.getBookSeries().getLock();
		lock.writeLock().lock();
		
		workbook = new HSSFWorkbook();
		
		for(NSheet sheet : book.getSheets()) {
			exportSheet(sheet);
		}
		
		exportNamedRange(book);
		
		try{
			
		} finally{
			lock.writeLock().unlock();
		}
	}

}
