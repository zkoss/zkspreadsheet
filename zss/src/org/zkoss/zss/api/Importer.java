package org.zkoss.zss.api;

import java.io.IOException;
import java.io.InputStream;

import org.zkoss.zss.api.model.Book;

public interface Importer {
		
	public Book imports(InputStream is, String bookName) throws IOException;
	
}
