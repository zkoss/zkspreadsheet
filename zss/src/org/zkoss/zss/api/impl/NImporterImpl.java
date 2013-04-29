package org.zkoss.zss.api.impl;

import java.io.InputStream;

import org.zkoss.zss.api.NImporter;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.impl.NBookImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XImporter;


public class NImporterImpl implements NImporter{
	XImporter importer;
	public NImporterImpl(XImporter importer) {
		this.importer = importer;
	}

	
	public NBook imports(InputStream is, String bookName){
		return new NBookImpl(new SimpleRef<XBook>(importer.imports(is, bookName)));
	}
	
}
