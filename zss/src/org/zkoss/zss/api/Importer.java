package org.zkoss.zss.api;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

import org.zkoss.zss.api.model.Book;

public interface Importer {
		
	public Book imports(InputStream is, String bookName) throws IOException;
	
	public Book imports(File file, String bookName) throws IOException;
	
	public Book imports(URL url, String bookName) throws IOException;
	
}
