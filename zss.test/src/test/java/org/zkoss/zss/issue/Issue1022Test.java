package org.zkoss.zss.issue;

import java.util.Locale;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zul.Filedownload;

public class Issue1022Test {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}

	/**
	 * Test a book with theme table.
	 * @throws IOException 
	 */
	@Test
	public void testImportThemeTable() throws IOException {
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try  {
			Book book = Util.loadBook(this, "book/1022-theme-table.xlsx");
			Assert.assertTrue("No Exception", true);
			Exporters.getExporter("xlsx").export(book, os);
			
			Sheet sheet1 = book.getSheetAt(0);
			Range rngC34 = Ranges.range(sheet1, "C34");
			CellData cd = rngC34.getCellData();
			Assert.assertEquals("Numeric", CellType.NUMERIC, cd.getResultType());
			Assert.assertEquals("C34", 23216d, cd.getDoubleValue(), 0d);
			
			Range rngF29 = Ranges.range(sheet1, "F29");
			cd = rngF29.getCellData();
			Assert.assertEquals("Numeric", CellType.NUMERIC, cd.getResultType());
			Assert.assertEquals("F29", 5952d, cd.getDoubleValue(), 0d);
			
		} catch (Exception e) {
			e.printStackTrace();
			Assert.assertTrue("Exception when load \"issue/book/1022-theme-table.xlsx\":\n" + e, false);
		} finally {
			 os.close();
		}
	}
}
