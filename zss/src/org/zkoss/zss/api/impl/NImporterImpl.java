package org.zkoss.zss.api.impl;

import java.io.InputStream;

import org.zkoss.zss.api.NImporter;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.impl.NBookImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.sys.Book;
import org.zkoss.zss.model.sys.Importer;


public class NImporterImpl implements NImporter{
	Importer importer;
	public NImporterImpl(Importer importer) {
		this.importer = importer;
	}

	
	public NBook imports(InputStream is, String bookName){
		return new NBookImpl(new SimpleRef<Book>(importer.imports(is, bookName)));
	}
	
}
