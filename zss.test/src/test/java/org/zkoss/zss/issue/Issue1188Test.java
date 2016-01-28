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
import org.zkoss.zss.model.STableStyle;
import org.zkoss.zss.model.STableStyleElem;
import org.zkoss.zss.model.impl.AbstractFillAdv;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.api.Exporters;

public class Issue1188Test {
	
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
			Book book0 = Util.loadBook(this, "book/1188-export-custom-table.xlsx");
			Assert.assertTrue("No Exception", true);
			Exporters.getExporter("xlsx").export(book0, os);
			
			is = new ByteArrayInputStream(os.toByteArray());
			String bookName = book0.getBookName();
			Book book = Importers.getImporter().imports(is, bookName);

			Assert.assertEquals("count", 1, book.getInternalBook().getTableStyles().size());
			
			STableStyle  tbStyle = book.getInternalBook().getTableStyles().get(0);
			Assert.assertEquals("name", "資產負債表", tbStyle.getName());
			
			STableStyleElem wholeTable = tbStyle.getWholeTableStyle();
			Assert.assertNotNull("wholeTable", wholeTable);
			
			STableStyleElem colStripe1 = tbStyle.getColStripe1Style();
			Assert.assertNotNull("firstColumnStripe", colStripe1);
			Assert.assertEquals("firstColumnStripe.size", 1, tbStyle.getColStripe1Size());
			
			STableStyleElem colStripe2 = tbStyle.getColStripe2Style();
			Assert.assertNull("secondColumnStripe", colStripe2);
			
			STableStyleElem rowStripe1 = tbStyle.getRowStripe1Style();
			Assert.assertNull("firstRowStripe", rowStripe1);
			
			STableStyleElem rowStripe2 = tbStyle.getRowStripe2Style();
			Assert.assertNull("SecondRowStripe", rowStripe2);
			
			STableStyleElem lastColumn = tbStyle.getLastColumnStyle();
			Assert.assertNull("lastColumn", lastColumn);
			
			STableStyleElem firstColumn = tbStyle.getFirstColumnStyle();
			Assert.assertNull("firstColumn", firstColumn);
			
			STableStyleElem headerRow = tbStyle.getHeaderRowStyle();
			Assert.assertNull("headerRow", headerRow);
			
			STableStyleElem totalRow = tbStyle.getTotalRowStyle();
			Assert.assertNotNull("totalRow", totalRow);
			
			STableStyleElem firstHeaderCell = tbStyle.getFirstHeaderCellStyle();
			Assert.assertNull("firstHeaderCell", firstHeaderCell);

			STableStyleElem lastHeaderCell = tbStyle.getLastHeaderCellStyle();
			Assert.assertNull("lastHeaderCell", lastHeaderCell);

			STableStyleElem firstTotalCell = tbStyle.getFirstTotalCellStyle();
			Assert.assertNull("firstTotalCell", firstTotalCell);
			
			STableStyleElem lastTotalCell = tbStyle.getLastTotalCellStyle();
			Assert.assertNull("lastTotalCell", lastTotalCell);
			
		} catch (Exception e) {
			e.printStackTrace();
			Assert.assertTrue("Exception when load \"issue/book/1188-export-custom-table.xlsx\":\n" + e, false);
		} finally {
			 os.close();
		}
	}
}
