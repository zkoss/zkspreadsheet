package org.zkoss.zss.api;

import java.io.FileOutputStream;

import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.model.Exporter;

public class NExporter {
	Exporter exporter;
	public NExporter(Exporter exporter){
		if(exporter==null){
			throw new IllegalAccessError("exporter not found");
		}
		
		this.exporter =exporter;
	}
	public void export(NBook book, FileOutputStream fos) {
		exporter.export(book.getNative(), fos);
	}
}
