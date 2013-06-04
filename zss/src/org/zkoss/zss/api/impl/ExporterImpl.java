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

import java.io.IOException;
import java.io.OutputStream;

import org.zkoss.poi.hssf.util.AreaReference;
import org.zkoss.poi.hssf.util.CellReference;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.BookImpl;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.sys.XExporter;
import org.zkoss.zss.model.sys.impl.Headings;
import org.zkoss.zss.ui.Rect;

/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class ExporterImpl implements Exporter {
	XExporter exporter;
	public ExporterImpl(XExporter exporter){
		if(exporter==null){
			throw new IllegalAccessError("exporter not found");
		}
		
		this.exporter =exporter;
	}
	public void export(Book book, OutputStream fos) throws IOException{
		exporter.export(((BookImpl)book).getNative(), fos);
	}
	public void export(Sheet sheet, OutputStream fos) throws IOException{
		exporter.export(((SheetImpl)sheet).getNative(), fos);
	}
	public void export(Sheet sheet,Rect selection,OutputStream fos) throws IOException{
		AreaReference af = new AreaReference(new CellReference(selection.getTop(), selection.getLeft()),
				new CellReference(selection.getBottom(), selection.getRight()));
		exporter.exportSelection(((SheetImpl)sheet).getNative(), af, fos);
	}
	@Override
	public boolean isSupportHeadings() {
		return exporter instanceof Headings;
	}
	@Override
	public void enableHeadings(boolean enable) {
		if(isSupportHeadings()){
			((Headings)exporter).enableHeadings(enable);
		}else{
			throw new RuntimeException("this export doesn't support headings");
		}
	}
}
