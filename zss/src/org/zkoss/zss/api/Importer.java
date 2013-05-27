package org.zkoss.zss.api;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

import org.zkoss.zss.api.model.Book;

/**
 * Importer for book
 * @author dennis
 * @since 3.0.0
 */
public interface Importer {
		
	/**
	 * Import book from a input stream
	 * @param is the input stream
	 * @param bookName the book name for imported book
	 * @return the book instance
	 * @throws IOException
	 */
	public Book imports(InputStream is, String bookName) throws IOException;
	
	/**
	 * Import book from a file
	 * @param file the file
	 * @param bookName the book name for imported book
	 * @return the book instance
	 * @throws IOException
	 */
	public Book imports(File file, String bookName) throws IOException;
	
	/**
	 * Import book from a URL
	 * @param url the url
	 * @param bookName the book name for imported book
	 * @return the book instance
	 * @throws IOException
	 */
	public Book imports(URL url, String bookName) throws IOException;
	
}
