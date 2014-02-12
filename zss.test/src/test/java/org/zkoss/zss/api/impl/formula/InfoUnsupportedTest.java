package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

@Ignore
public class InfoUnsupportedTest extends FormulaTestBase {
	
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
	
	// #NAME?
	@Test
	public void testCELL()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testCELL(book);
	}
	
	// #NAME?
	@Test
	public void testINFO()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testINFO(book);
	}
	
}
