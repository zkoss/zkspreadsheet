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
import org.zkoss.zss.model.SDataValidation.ValidationType;
import org.zkoss.zss.model.impl.AbstractSheetAdv;
import org.zkoss.zss.model.impl.DataValidationImpl;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.api.Exporters;

public class Issue1147Test {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.GERMANY);
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
	public void testDataValidationNumber() throws IOException {
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try  {
			Book book = Util.loadBook(this, "book/blank.xlsx");
			Assert.assertTrue("No Exception", true);
			AbstractSheetAdv sheet = (AbstractSheetAdv) book.getSheetAt(0).getInternalSheet();
			DataValidationImpl dv= new DataValidationImpl(sheet, "dummy");
			dv.setValidationType(ValidationType.INTEGER);
			dv.setFormulas("1", "111");
			Assert.assertEquals("formula1", "1", dv.getFormula1());
			Assert.assertEquals("formula2", "111", dv.getFormula2());
			dv.setValidationType(ValidationType.DECIMAL);
			dv.setFormulas("1,5", "123,4");
			Assert.assertEquals("formula1", "1,5", dv.getFormula1());
			Assert.assertEquals("formula2", "123,4", dv.getFormula2());
		} catch (Exception e) {
			Assert.assertTrue("Exception when load \"issue/book/1009-import-NPE.xlsx\":\n" + e, false);
		} finally {
			 os.close();
		}
	}
}
