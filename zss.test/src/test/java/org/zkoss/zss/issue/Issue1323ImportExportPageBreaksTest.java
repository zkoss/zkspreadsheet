package org.zkoss.zss.issue;

import java.util.Locale;
import java.io.ByteArrayInputStream;
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
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SCustomFilter;
import org.zkoss.zss.model.SCustomFilters;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.api.Exporters;

public class Issue1323ImportExportPageBreaksTest {
	
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
	 * Test a book with print area and export it and it should not throw exception
	 * @throws IOException 
	 */
	@Test
	public void testImportExportCustomFilters() throws IOException {
		ByteArrayInputStream is = null;
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try  {
			Book book = Util.loadBook(this, "book/1323-page-breaks.xlsx");
			Sheet sheet1 = book.getSheetAt(0);
			SSheet ssheet1 = sheet1.getInternalSheet();
			validPageBreaks(ssheet1);
			Assert.assertTrue("No Exception", true);
			try {
				Exporters.getExporter("xlsx").export(book, os);
			} finally {
				os.close();
			}
			is = new ByteArrayInputStream(os.toByteArray());
			Book book2 = Util.loadBook(is, "1323-page-breaks.xlsx");
			Sheet sheet2 = book2.getSheetAt(0);
			SSheet ssheet2 = sheet2.getInternalSheet();
			validPageBreaks(ssheet2);
		} catch (Exception e) {
			e.printStackTrace();
			Assert.assertTrue("Exception when import or export \"issue/book/1224-custom-filters.xlsx\":\n" + e, false);
		} finally {
			is.close();
		}
	}
	
	private void validPageBreaks(SSheet ssheet) {
		Assert.assertEquals("size", 1, ssheet.getViewInfo().getRowBreaks().length);
		Assert.assertEquals("rowIndex", 4, ssheet.getViewInfo().getRowBreaks()[0]);
		Assert.assertEquals("size", 1, ssheet.getViewInfo().getColumnBreaks().length);
		Assert.assertEquals("rowIndex", 2, ssheet.getViewInfo().getColumnBreaks()[0]);
	}
}
