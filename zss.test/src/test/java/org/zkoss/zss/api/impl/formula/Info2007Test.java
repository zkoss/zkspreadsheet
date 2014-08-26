package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Info2007Test extends FormulaTestBase {
	
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
	
	// ERROR.TYPE
	@Test
	public void testERRORTYPE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testERRORTYPE(book);
	}

	// ISBLANK
	@Test
	public void testISBLANK() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISBLANK(book);
	}

	// ISERR
	@Test
	public void testISERR() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISERR(book);
	}

	// ISERROR
	@Test
	public void testISERROR() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISERROR(book);
	}

	// ISEVEN
	@Test
	public void testISEVEN() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISEVEN(book);
	}

	// ISLOGICAL
	@Test
	public void testISLOGICAL() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISLOGICAL(book);
	}

	// ISNA
	@Test
	public void testISNA() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISNA(book);
	}

	// ISNONTEXT
	@Test
	public void testISNONTEXT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISNONTEXT(book);
	}

	// ISNUMBER
	@Test
	public void testISNUMBER() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISNUMBER(book);
	}

	// ISODD
	@Test
	public void testISODD() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISODD(book);
	}

	// ISREF
	@Test
	public void testISREF() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISREF(book);
	}

	// ISTEXT
	@Test
	public void testISTEXT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testISTEXT(book);
	}
	
	// N
	@Test
	public void testN() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testN(book);
	}

	// NA
	@Test
	public void testNA() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testNA(book);
	}

	// TYPE
	@Test
	public void testTYPE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testTYPE(book);
	}
}
