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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.zkoss.poi.POIXMLDocument;
import org.zkoss.poi.poifs.filesystem.POIFSFileSystem;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.CellRef;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XImporter;
import org.zkoss.zss.model.sys.XModelException;
import org.zkoss.zss.model.sys.XSheet;

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
			return postProcess(new HSSFBookImpl(bookname, is));
		}
		if(POIXMLDocument.hasOOXMLHeader(is)) {
			return postProcess(new XSSFBookImpl(bookname, is));
		}
		throw new IllegalArgumentException("InputStream was neither an OLE2 stream, nor an OOXML stream");
	}
	
	private XBook postProcess(XBook book) {
		// ZSS-621: replace all shared formulas with normal formulas
		for(int i = 0; i < book.getNumberOfSheets(); ++i) {
			XSheet sheet = book.getWorksheetAt(i);
			List<AreaRef> references = sheet.getSharedFormulaReferences();
			replaceSharedFormula(sheet, references);
		}
		return book;
	}

	private void replaceSharedFormula(XSheet sheet, List<AreaRef> regions) {
		for (AreaRef ref : regions) {

			// read and buffer all formulas first then apply them at once
			// because shared formulas header can be any location of the region
			Map<CellRef, String> buf = new LinkedHashMap<CellRef, String>();
			for (int r = ref.getRow(); r <= ref.getLastRow(); ++r) {
				Row row = sheet.getRow(r);
				if (row != null) {
					for (int c = ref.getColumn(); c <= ref.getLastColumn(); ++c) {
						Cell cell = row.getCell(c);
						if (cell != null && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
							buf.put(new CellRef(r, c), cell.getCellFormula());
						}
					}
				}
			}
			for (Entry<CellRef, String> entry : buf.entrySet()) {
				CellRef cr = entry.getKey();
				Row row = sheet.getRow(cr.getRow());
				if (row != null) { // just in case
					Cell cell = row.getCell(cr.getColumn());
					if (cell != null) { // just in case
						BookHelper.setCellFormulaForImport(cell, entry.getValue());
					}
				}
			}
		}
	}
}
