/* ExcelImporter.java

	Purpose:
		
	Description:
		
	History:
		Mar 12, 2010 2:25:01 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.sys.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.net.URL;

import org.zkoss.poi.POIXMLDocument;
import org.zkoss.poi.poifs.filesystem.POIFSFileSystem;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XImporter;
import org.zkoss.zss.model.sys.XModelException;

/**
 * Imports an Excel file into as a {@link XBook}.
 * @author henrichen
 *
 */
public class ExcelImporter implements XImporter {
	@Override
	public XBook imports(String filename) throws IOException {
		InputStream is = null;
		try {
			is = new FileInputStream(filename);
			return importsFromStream(is, filename);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					throw XModelException.Aide.wrap(e);
				}
			}
		}
	}

	@Override
	public XBook imports(File file) throws IOException{
		final String name = file.getName();
		InputStream is = null;
		try {
			is = new FileInputStream(file);
			return importsFromStream(is, name);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					throw XModelException.Aide.wrap(e);
				}
			}
		}
	}

	@Override
	public XBook imports(InputStream is, String bookname) throws IOException{
		return importsFromStream(is, bookname);
	}
	
	public XBook importsFromURL(URL url) throws IOException{
		InputStream is = null;
		try {
			is = url.openStream();
			return importsFromStream(is, url.toString());
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch(IOException ex) {
					throw XModelException.Aide.wrap(ex);
				}
			}
		}
	}
	
	private XBook importsFromStream(InputStream is, String bookname) 
	throws IOException {
		final int j = bookname.lastIndexOf("/");
		if (j >=0) {
			bookname = bookname.substring(j+1);
		}
		
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
		throw new IllegalArgumentException("InputStream was neither an OLE2 stream, nor an OOXML stream");
	}
}
