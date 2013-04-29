package org.zkoss.zss.api.impl;

import java.io.FileOutputStream;

import org.zkoss.zss.api.NExporter;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.impl.NBookImpl;
import org.zkoss.zss.model.sys.Exporter;

public class NExporterImpl implements NExporter {
	Exporter exporter;
	public NExporterImpl(Exporter exporter){
		if(exporter==null){
			throw new IllegalAccessError("exporter not found");
		}
		
		this.exporter =exporter;
	}
	public void export(NBook book, FileOutputStream fos) {
		exporter.export(((NBookImpl)book).getNative(), fos);
	}
}
