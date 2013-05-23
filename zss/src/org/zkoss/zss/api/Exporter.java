package org.zkoss.zss.api;

import java.io.IOException;
import java.io.OutputStream;

import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;

public interface Exporter {
	
	public void export(Book book, OutputStream fos) throws IOException;
	public void export(Sheet sheet, OutputStream fos) throws IOException;
	public void export(Sheet sheet,Rect selection,OutputStream fos) throws IOException;

	public boolean isSupportHeadings();
	public void enableHeadings(boolean enable);
}
