package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Text2007Test extends FormulaTestBase {
	
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

	// CHAR
	@Test
	public void testCHAR() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCHAR(book);
	}
	
	// CLEAN
	@Test
	public void testCLEAN() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCLEAN(book);
	}

	// CODE
	@Test
	public void testCODE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCODE(book);
	}
	
	// CDONCATENATE
	@Test
	public void testCONCATENATE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCONCATENATE(book);
	}
	
	// DOLLAR
	@Test
	public void testDOLLAR()  {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testDOLLAR(book);
	}
	
	// EXACT
	@Test
	public void testEXACT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testEXACT(book);
	}

	// FIND
	@Test
	public void testFIND() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testFIND(book);
	}

	// FIXED
	@Test
	public void testFIXED() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testFIXED(book);
	}

	// LEFT
	@Test
	public void testLEFT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testLEFT(book);
	}

	// LEN
	@Test
	public void testLEN() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testLEN(book);
	}
	
	// LOWER
	@Test
	public void testLOWER() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testLOWER(book);
	}
	
	// MID
	@Test
	public void testMID() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testMID(book);
	}

	// PROPER
	@Test
	public void testPROPER() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testPROPER(book);
	}

	// REPLACE
	@Test
	public void testREPLACE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testREPLACE(book);
	}

	// REPT
	@Test
	public void testREPT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testREPT(book);
	}

	// RIGHT
	@Test
	public void testRIGHT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testRIGHT(book);
	}

	// SEARCH
	@Test
	public void testSEARCH() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testSEARCH(book);
	}

	// SUBSTITUTE
	@Test
	public void testSUBSTITUTE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testSUBSTITUTE(book);
	}

	// T
	@Test
	public void testT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testT(book);
	}

	// TEXT
	@Test
	public void testTEXT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testTEXT(book);
	}

	// TRIM
	@Test
	public void testTRIM() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testTRIM(book);
	}

	// UPPER
	@Test
	public void testUPPER() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testUPPER(book);
	}
	
	// VALUE
	@Test
	public void testVALUE()  {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testVALUE(book);
	}
}
