package org.zkoss.zss.issue;

import java.io.File;
import java.net.URL;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.model.ImExpTestUtil;
import org.zkoss.zss.model.ImporterTest;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SDataValidation.ValidationType;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.range.impl.imexp.ExcelExportFactory;

public class Issue835Test {
	
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

	protected URL DATAVALIDATTIONIMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/import.xlsx");
	@Test
	public void testDataValidationExport() {
		URL dvURL = getClass().getResource("book/835-validationType-any.xlsx");
		File outFile = ImExpTestUtil.write(ImExpTestUtil.loadBook(dvURL, "XSSFBook"), ExcelExportFactory.Type.XLSX);
		SBook book = ImExpTestUtil.loadBook(outFile, "PoiBook");
		dataValidationTest(book);
	}
	
	private void dataValidationTest(SBook book) {
		SSheet sheet = book.getSheet(0);
		SDataValidation dv = sheet.getDataValidation(8, 3); //"D9"
		Assert.assertEquals("ValidationType", ValidationType.ANY, dv.getValidationType());
	}
}
