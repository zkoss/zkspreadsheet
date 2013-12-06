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
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.zss.ngapi.NExporter;
import org.zkoss.zss.ngapi.impl.ExcelExportFactory.Type;
import org.zkoss.zss.ngmodel.NBook;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NExcelExporter extends AbstractExporter{
	
	private final ExcelExportFactory.Type type;
	
	public NExcelExporter(Type type){
		this.type = type;
	}
	
	@Override
	public void export(NBook book, File file) throws IOException {
		OutputStream os = null;
		try{
			os = new FileOutputStream(file);
			export(book,os);
		}finally{
			if(os!=null){
				try{
					os.close();
				}catch(Exception x){};
			}
		}
	}
	
	@Override
	public void export(NBook book, OutputStream fos) throws IOException {
		ReadWriteLock lock = book.getBookSeries().getLock();
		lock.writeLock().lock();
		try{
			//implement here
			
			
		}finally{
			lock.writeLock().unlock();
		}
	}

}
