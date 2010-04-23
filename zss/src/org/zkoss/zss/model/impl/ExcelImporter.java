/* ExcelImporter.java

	Purpose:
		
	Description:
		
	History:
		Mar 12, 2010 2:25:01 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.net.URL;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Importer;
import org.zkoss.zss.model.ModelException;

/**
 * Imports an Excel file into as a {@Book}.
 * @author henrichen
 *
 */
public class ExcelImporter implements Importer {
	@Override
	public Book imports(String filename) {
		InputStream is = null;
		try {
			is = new FileInputStream(filename);
			return importsFromStream(is, filename);
		} catch (Exception e) {
			throw ModelException.Aide.wrap(e);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					throw ModelException.Aide.wrap(e);
				}
			}
		}
	}

	@Override
	public Book imports(File file) {
		final String name = file.getName();
		InputStream is = null;
		try {
			is = new FileInputStream(file);
			return importsFromStream(is, name);
		} catch (Exception e) {
			throw ModelException.Aide.wrap(e);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					throw ModelException.Aide.wrap(e);
				}
			}
		}
	}

	@Override
	public Book imports(InputStream is, String bookname) {
		try {
			return importsFromStream(is, bookname);
		} catch (Exception e) {
			throw ModelException.Aide.wrap(e);
		}
	}
	
	public Book importsFromURL(URL url) {
		InputStream is = null;
		try {
			is = url.openStream();
			return importsFromStream(is, url.toString());
		} catch (Exception ex) {
			throw ModelException.Aide.wrap(ex);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch(IOException ex) {
					throw ModelException.Aide.wrap(ex);
				}
			}
		}
	}
	
	private Book importsFromStream(InputStream is, String bookname) 
	throws IOException {
		final int j = bookname.lastIndexOf("/");
		bookname = bookname.substring(j+1);
		
		// If inputstream doesn't do mark/reset, wrap up
		if(!is.markSupported()) {
			is = new PushbackInputStream(is, 8);
		}
		if(POIFSFileSystem.hasPOIFSHeader(is)) {
			return new HSSFBookImpl(bookname, is);
		}
		if(POIXMLDocument.hasOOXMLHeader(is)) {
			return new XSSFBookImpl(bookname, is);
		}
		throw new IllegalArgumentException("Your InputStream was neither an OLE2 stream, nor an OOXML stream");
	}
}
