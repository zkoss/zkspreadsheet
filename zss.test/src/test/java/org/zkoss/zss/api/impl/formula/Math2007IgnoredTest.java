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
public class Math2007IgnoredTest extends FormulaTestBase {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssContextLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssContextLocale();
	}
	
	// This should be check by human
	@Test
	public void testRANDBETWEEN()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testRANDBETWEEN(book);
	}
	
	// This should be check by human
	@Test
	public void testRAND()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testRAND(book);
	}
}
