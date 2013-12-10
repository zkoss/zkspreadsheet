package org.zkoss.zss.ngmodel;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.io.InputStream;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.impl.NExcelImportAdapter;

public class ImporterTest {
	
	@Test
	public void loadXSSFBook() {
		final InputStream is = ImporterTest.class.getResourceAsStream("book/import.xlsx");
		NImporter importer = new NExcelImportAdapter();
		NBook book = null;
		try {
			book = importer.imports(is, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		if(book != null) {
			assertEquals(book.getBookName(), "XSSFBook");
			
			// 3 sheet
			assertEquals(book.getNumOfSheet(), 3);
			
			NSheet sheet1 = book.getSheet(0);
			assertEquals(sheet1.getSheetName(), "First");
			NSheet sheet2 = book.getSheet(1);
			assertEquals(sheet2.getSheetName(), "Second");
			NSheet sheet3 = book.getSheet(2);
			assertEquals(sheet3.getSheetName(), "Third");
			
		} else {
			Assert.fail("book is null!");
		}
	}


}
