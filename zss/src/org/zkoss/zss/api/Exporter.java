package org.zkoss.zss.api;

import java.io.FileOutputStream;
import java.io.IOException;

import org.zkoss.zss.api.model.Book;

public interface Exporter {
	
	public void export(Book book, FileOutputStream fos) throws IOException;
}
