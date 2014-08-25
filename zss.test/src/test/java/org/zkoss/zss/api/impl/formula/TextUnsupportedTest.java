package org.zkoss.zss.api.impl.formula;

import static org.junit.Assert.assertEquals;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

@Ignore
public class TextUnsupportedTest extends FormulaTestBase {
	
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
	public void testASC()  {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testASC(book);
	}
	
	// #NAME?
	protected void testFINDB(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("1", Ranges.range(sheet, "B21").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testLEFTB()  {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testLEFTB(book);
	}
	
	// #NAME?
	@Test
	public void testLENB()  {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testLENB(book);
	}
	
	// #NAME?
	@Test
	public void testMIDB()  {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testMIDB(book);
	}
	
	// #NAME?
	@Test
	public void testSEARCHB()  {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testSEARCHB(book);
	}
	
	// #NAME?
	@Test
	public void testREPLACEB()  {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testREPLACEB(book);
	}
	
	// #NAME?
	@Test
	public void testRIGHTB()  {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testRIGHTB(book);
	}

}
