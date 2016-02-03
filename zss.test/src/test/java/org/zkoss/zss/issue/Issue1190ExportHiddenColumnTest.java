package org.zkoss.zss.issue;


import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.junit.Assert;
import org.junit.Test;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SColumn;

public class Issue1190ExportHiddenColumnTest {

	@Test
	public void testHiddenAfterSaveReload() throws IOException {
		Book book = Util.loadBook(this, "book/1190-export-hidden-column.xlsx");
		Sheet sheet = book.getSheetAt(0);
		Range c = Ranges.range(sheet, "A:A").toColumnRange(); 
		c.setHidden(true); // hide column A

		SColumn firstColumn = sheet.getInternalSheet().getColumn(0);

		//at this point the first column is hidden successfully
		Assert.assertTrue("hidden first column not set successfully", firstColumn.isHidden());
		
		//saving the document to buffer
		ByteArrayOutputStream bufferOut = new ByteArrayOutputStream();
		save(book, bufferOut);
		
		//reloading the document from buffer
		Book reloadedBook = load(new ByteArrayInputStream(bufferOut.toByteArray()));
		SColumn reloadedFirstColumn = reloadedBook.getSheetAt(0).getInternalSheet().getColumn(0);
		//after reloading the first column is no longer hidden
		Assert.assertTrue("hidden first column not saved/reloaded successfully", reloadedFirstColumn.isHidden());
	}

	public Book load(InputStream is) throws IOException {
		Importer importer = Importers.getImporter();
		Book book = importer.imports(is, "template");
		return book;
	}
	
	public void save(Book book, OutputStream out) throws IOException { 
		Exporters.getExporter("excel").export(book, out);
	}

	
}
