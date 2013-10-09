package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Info2003Test extends FormulaTestBase {
	
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
	
	// ERROR.TYPE
	@Test
	public void testERRORTYPE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testERRORTYPE(book);
	}

	// ISBLANK
	@Test
	public void testISBLANK() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISBLANK(book);
	}

	// ISERR
	@Test
	public void testISERR() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISERR(book);
	}

	// ISERROR
	@Test
	public void testISERROR() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISERROR(book);
	}

	// ISEVEN
	@Test
	public void testISEVEN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISEVEN(book);
	}

	// ISLOGICAL
	@Test
	public void testISLOGICAL() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISLOGICAL(book);
	}

	// ISNA
	@Test
	public void testISNA() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISNA(book);
	}

	// ISNONTEXT
	@Test
	public void testISNONTEXT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISNONTEXT(book);
	}

	// ISNUMBER
	@Test
	public void testISNUMBER() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISNUMBER(book);
	}

	// ISODD
	@Test
	public void testISODD() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISODD(book);
	}

	// ISREF
	@Test
	public void testISREF() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISREF(book);
	}

	// ISTEXT
	@Test
	public void testISTEXT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testISTEXT(book);
	}
	
	// N
	@Test
	public void testN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testN(book);
	}

	// NA
	@Test
	public void testNA() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testNA(book);
	}

	// TYPE
	@Test
	public void testTYPE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testTYPE(book);
	}
}
