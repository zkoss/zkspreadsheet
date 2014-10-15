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
import org.zkoss.zss.model.SName;

public class Issue803Test {
	
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
	public void testFormulaName() throws IOException {
		Book book = Util.loadBook(this, "book/803-formula-name.xlsx");
		
		String formulaName = "lstYears";
		SName name = book.getInternalBook().getNameByName(formulaName);
		String formulaString = name.getRefersToFormula();
		Assert.assertEquals("lstYears", "OFFSET('Financial Data Input'!$B$5:$I$5,0,1,1,COUNTA('Financial Data Input'!$B$5:$I$5)-1)", formulaString);
		
		//export and import then test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		name = book.getInternalBook().getNameByName(formulaName);
		formulaString = name.getRefersToFormula();
		Assert.assertEquals("lstYears", "OFFSET('Financial Data Input'!$B$5:$I$5,0,1,1,COUNTA('Financial Data Input'!$B$5:$I$5)-1)", formulaString);
	}
}