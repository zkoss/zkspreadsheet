package org.zkoss.zss.ngapi;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

import org.zkoss.zss.ngmodel.NBook;


public interface NExporter {
	/**
	 * Export book
	 * @param book the book to export
	 * @param fos the output stream to store data
	 * @throws IOException
	 */
	public void export(NBook book, OutputStream fos) throws IOException;
	
	/**
	 * Export book
	 * @param book the book to export
	 * @param fos the output file to store data
	 * @throws IOException
	 */
	public void export(NBook book, File file) throws IOException;
}
