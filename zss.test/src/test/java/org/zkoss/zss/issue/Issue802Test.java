package org.zkoss.zss.issue;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

public class Issue802Test {
	
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
	 * parse a defined name which contains question mark '?'
	 * @throws IOException 
	 */
	@Test
	public void exportSheetWithPrintArea() throws IOException {
		Book book = Util.loadBook(this, "book/802-print-area-export.xlsx");
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		Sheet sheet1 = book.getSheetAt(0);
		String printArea = sheet1.getInternalSheet().getPrintSetup().getPrintArea();
		Assert.assertEquals("Print Area", "Sheet1!$B$2:$D$6", printArea);
	}
}