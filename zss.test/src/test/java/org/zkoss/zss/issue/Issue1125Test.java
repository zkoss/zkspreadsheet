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

public class Issue1125Test {
	
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
	public void testGetMergedRegion() throws IOException {
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try  {
			Book book = Util.loadBook(this, "book/1125-get-mergedRegion.xlsx");
			Assert.assertTrue("No Exception", true);
			Exporters.getExporter("xlsx").export(book, os);
			
			Sheet sheet1 = book.getSheetAt(0);
			Range A1 = Ranges.range(sheet1, "A1");
			Range B1 = Ranges.range(sheet1, "B1");
			Range A2 = Ranges.range(sheet1, "A2");
			Range B2 = Ranges.range(sheet1, "B2");
			Range C1 = Ranges.range(sheet1, "C1");
			CellRegion rgn = new CellRegion("A1:B2");
			
			Assert.assertEquals("A1-merge", rgn, A1.getMergedRegion());
			Assert.assertEquals("B1-merge", rgn, B1.getMergedRegion());
			Assert.assertEquals("A2-merge", rgn, A2.getMergedRegion());
			Assert.assertEquals("B2-merge", rgn, B2.getMergedRegion());
			
			Assert.assertNull("C1-merge", C1.getMergedRegion());
			
		} catch (Exception e) {
			e.printStackTrace();
			Assert.assertTrue("Exception when load \"issue/book/1022-theme-table.xlsx\":\n" + e, false);
		} finally {
			 os.close();
		}
	}
}
