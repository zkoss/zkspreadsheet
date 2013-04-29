package org.zkoss.zss.api.impl;

import java.io.FileOutputStream;

import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.impl.BookImpl;
import org.zkoss.zss.model.sys.XExporter;

public class ExporterImpl implements Exporter {
	XExporter exporter;
	public ExporterImpl(XExporter exporter){
		if(exporter==null){
			throw new IllegalAccessError("exporter not found");
		}
		
		this.exporter =exporter;
	}
	public void export(Book book, FileOutputStream fos) {
		exporter.export(((BookImpl)book).getNative(), fos);
	}
}
