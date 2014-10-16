package org.zkoss.zss.issue;

import java.io.IOException;
import java.util.Locale;

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
import org.zkoss.zss.api.model.Sheet;

public class Issue804Test {
	
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
	public void evalNamedFormula() throws IOException {
		Book book = Util.loadBook(this, "book/804-name-formula-range.xlsx");
		Sheet sheet1 = book.getSheetAt(0);
		Range rngA1 = Ranges.range(sheet1, "A1");
		CellData a1 = rngA1.getCellData();
		Assert.assertEquals("Formula", "OFFSET(lstYears,0,1,1,1)", a1.getFormulaValue());
		Assert.assertEquals("Value a Number Type", CellData.CellType.NUMERIC, a1.getResultType());
		Assert.assertEquals("Value", 2009d, a1.getDoubleValue(), 0d);
		
//		//export and test again
//		File temp = Setup.getTempFile();
//		Exporters.getExporter().export(book, temp);
//		book = Importers.getImporter().imports(temp, "test");
//		Sheet sheet1 = book.getSheetAt(0);
//		String printArea = sheet1.getInternalSheet().getPrintSetup().getPrintArea();
//		Assert.assertEquals("Print Area", "Sheet1!$B$2:$D$6", printArea);
	}
}