package org.zkoss.zss.api.impl;

import java.io.InputStream;

import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.impl.BookImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XImporter;


public class ImporterImpl implements Importer{
	XImporter importer;
	public ImporterImpl(XImporter importer) {
		this.importer = importer;
	}

	
	public Book imports(InputStream is, String bookName){
		return new BookImpl(new SimpleRef<XBook>(importer.imports(is, bookName)));
	}


	public XImporter getNative() {
		return importer;
	}
	
}
