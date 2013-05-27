package org.zkoss.zss.api;

import java.io.IOException;
import java.io.OutputStream;

import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;

/**
 * Exporter for book, sheet
 * @author dennis
 * @since 3.0.0
 */
public interface Exporter {
	
	/**
	 * Export book
	 * @param book the book to export
	 * @param fos the output stream to store data
	 * @throws IOException
	 */
	public void export(Book book, OutputStream fos) throws IOException;
	/**
	 * Export sheet
	 * @param sheet the sheet to export
	 * @param fos the output stream to store data
	 * @throws IOException
	 */
	public void export(Sheet sheet, OutputStream fos) throws IOException;
	/**
	 * Export selection of sheet
	 * @param sheet the sheet to export
	 * @param selection the selection to export
	 * @param fos the output stream to store data
	 * @throws IOException
	 */
	public void export(Sheet sheet,Rect selection,OutputStream fos) throws IOException;

	/**
	 * @return true if this exporter support heading configuration
	 */
	public boolean isSupportHeadings();
	
	/**
	 * Sets heading configuration,
	 * @param enable true to enable heading
	 */
	public void enableHeadings(boolean enable);
}
