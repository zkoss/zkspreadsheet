package org.zkoss.zss.api;

import java.io.InputStream;

import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.model.Importer;

public class NImporter {
	Importer importer;
	public NImporter(Importer importer) {
		this.importer = importer;
	}

	
	public NBook imports(InputStream is, String bookName){
		return new NBook(importer.imports(is, bookName));
	}
	
}
