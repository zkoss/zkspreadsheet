/* ExporterImpl.java

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
package org.zkoss.zss.api.impl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

//import org.zkoss.poi.hssf.util.AreaReference;
//import org.zkoss.poi.hssf.util.CellReference;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.BookImpl;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.sys.XExporter;
import org.zkoss.zss.model.sys.impl.Headings;

/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class ExporterImpl implements Exporter {
	private XExporter _exporter;
	public ExporterImpl(XExporter exporter){
		if(exporter==null){
			throw new IllegalAccessError("exporter not found");
		}
		
		this._exporter =exporter;
	}
	public void export(Book book, OutputStream fos) throws IOException{
		_exporter.export(((BookImpl)book).getNative(), fos);
	}
	public void export(Book book, File file) throws IOException{
		FileOutputStream fos = null;
		try{
			fos = new FileOutputStream(file);
			_exporter.export(((BookImpl)book).getNative(), fos);
		}finally{
			if(fos!=null){
				try{
					fos.close();
				}catch(Exception x){}//eat
			}
		}
	}
	public void export(Sheet sheet, OutputStream fos) throws IOException{
		_exporter.export(((SheetImpl)sheet).getNative(), fos);
	}
	public void export(Sheet sheet,AreaRef selection,OutputStream fos) throws IOException{
		org.zkoss.poi.ss.util.AreaReference af = new org.zkoss.poi.ss.util.AreaReference(new org.zkoss.poi.ss.util.CellReference(selection.getRow(), selection.getColumn()),
				new org.zkoss.poi.ss.util.CellReference(selection.getLastRow(), selection.getLastColumn()));
		_exporter.exportSelection(((SheetImpl)sheet).getNative(), af, fos);
	}
//	@Override
//	public boolean isSupportHeadings() {
//		return exporter instanceof Headings;
//	}
//	@Override
//	public void enableHeadings(boolean enable) {
//		if(isSupportHeadings()){
//			((Headings)exporter).enableHeadings(enable);
//		}else{
//			throw new RuntimeException("this export doesn't support headings");
//		}
//	}
}
