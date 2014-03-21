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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import org.zkoss.poi.POIXMLDocument;
import org.zkoss.poi.poifs.filesystem.POIFSFileSystem;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Importer;
import org.zkoss.zss.model.ModelException;
import org.zkoss.zss.model.Worksheet;

/**
 * Imports an Excel file into as a {@link Book}.
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
		throw new IllegalArgumentException("Your InputStream was neither an OLE2 stream, nor an OOXML stream");
	}
	
	private Book postProcess(Book book) {
		// ZSS-621: replace all shared formulas with normal formulas
		for(int i = 0; i < book.getNumberOfSheets(); ++i) {
			Worksheet sheet = book.getWorksheetAt(i);
			List<AreaReference> references = sheet.getSharedFormulaReferences();
			replaceSharedFormula(sheet, references);
		}
		return book;
	}

	private void replaceSharedFormula(Worksheet sheet, List<AreaReference> regions) {
		for(AreaReference ref : regions) {

			// read and buffer all formulas first then apply them at once
			// because shared formulas header can be any location of the region
			Map<CellReference, String> buf = new LinkedHashMap<CellReference, String>();
			for(CellReference cr : ref.getAllReferencedCells()) {
				Row row = sheet.getRow(cr.getRow());
				if(row != null) {
					Cell cell = row.getCell(cr.getCol());
					if(cell != null && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
						buf.put(cr, cell.getCellFormula());
					}
				}
			}
			for(Entry<CellReference, String> entry : buf.entrySet()) {
				CellReference cr = entry.getKey();
				Row row = sheet.getRow(cr.getRow());
				if(row != null) { // just in case
					Cell cell = row.getCell(cr.getCol());
					if(cell != null) { // just in case
						BookHelper.setCellFormulaForImport(cell, entry.getValue());
					}
				}
			}
		}
	}
}
