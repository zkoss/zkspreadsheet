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
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zul.Filedownload;

public class Issue1124Test {
	
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
	public void testGetDataRegion() throws IOException {
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try  {
			Book book = Util.loadBook(this, "book/1124-get-dataRegion.xlsx");
			Assert.assertTrue("No Exception", true);
			Exporters.getExporter("xlsx").export(book, os);
			
			Sheet sheet1 = book.getSheetAt(0);
			Sheet sheet2 = book.getSheetAt(1);
			Sheet sheet3 = book.getSheetAt(2);
			
			Range rngSheet1 = Ranges.range(sheet1);
			Range rngSheet2 = Ranges.range(sheet2);
			Range rngSheet3 = Ranges.range(sheet3);
			
			CellRegion rgn1 = rngSheet1.getDataRegion();
			CellRegion rgn2 = rngSheet2.getDataRegion();
			CellRegion rgn3 = rngSheet3.getDataRegion();
			
			CellRegion expectRgn1 = new CellRegion("A1:XFD1");
			CellRegion expectRgn2 = new CellRegion("A1:A1");
			
			Assert.assertEquals("sheet1-data", expectRgn1, rgn1);
			Assert.assertEquals("sheet2-data", expectRgn2, rgn2);			
			Assert.assertNull("sheet3-data", rgn3);
			
		} catch (Exception e) {
			e.printStackTrace();
			Assert.assertTrue("Exception when load \"issue/book/1124-get-dataRegion.xlsx\":\n" + e, false);
		} finally {
			 os.close();
		}
	}
}
