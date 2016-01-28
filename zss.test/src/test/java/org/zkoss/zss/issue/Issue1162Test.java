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
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SExtraStyle;
import org.zkoss.zss.model.SFill.FillPattern;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.impl.AbstractFillAdv;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.api.Exporters;

public class Issue1162Test {
	
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
	public void testExportXlsBookWithPrintArea() throws IOException {
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		ByteArrayInputStream is;
		try  {
			Book book0 = Util.loadBook(this, "book/1162-export-condition-format.xlsx");
			Assert.assertTrue("No Exception", true);
			Exporters.getExporter("xlsx").export(book0, os);
			
			is = new ByteArrayInputStream(os.toByteArray());
			String bookName = book0.getBookName();
			Book book = Importers.getImporter().imports(is, bookName);

			SExtraStyle  extraStyle = book.getInternalBook().getExtraStyleAt(0);
			SColor bgcolor = extraStyle.getBackColor();
			Assert.assertEquals("bgColor", "#ffc7ce", bgcolor.getHtmlColor());
			AbstractFillAdv fill = (AbstractFillAdv)extraStyle.getFill();
			FillPattern fillPattern = fill.getRawFillPattern();
			Assert.assertNull("fillPattern", fillPattern);
//			Filedownload.save(os.toByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "scale0.xlsx"); 
		} catch (Exception e) {
			e.printStackTrace();
			Assert.assertTrue("Exception when load \"issue/book/1162-export-condition-format.xlsx\":\n" + e, false);
		} finally {
			 os.close();
		}
	}
}
