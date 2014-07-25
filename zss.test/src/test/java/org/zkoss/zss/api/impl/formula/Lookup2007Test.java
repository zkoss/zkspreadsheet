package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Lookup2007Test extends FormulaTestBase {
	
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
	
	// ADDRESS
	@Test
	public void testADDRESS() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testADDRESS(book);
	}

	// CHOOSE
	@Test
	public void testCHOOSE() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCHOOSE(book);
	}

	// COLUMN
	@Test
	public void testCOLUMN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCOLUMN(book);
	}

	// COLUMNS
	@Test
	public void testCOLUMNS() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCOLUMNS(book);
	}

	// HLOOKUP
	@Test
	public void testHLOOKUP() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testHLOOKUP(book);
	}

	// HYPERLINK
	@Test
	public void testHyperLink() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testHyperLink(book);
	}
	
	// INDEX
	@Test
	public void testINDEX()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testINDEX(book);
	}

	// INDIRECT
	@Test
	public void testINDIRECT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testINDIRECT(book);
	}

	// LOOKUP
	@Test
	public void testLOOKUP() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testLOOKUP(book);
	}

	// MATCH
	@Test
	public void testMATCH() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMATCH(book);
	}

	// OFFSET
	@Test
	public void testOFFSET() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testOFFSET(book);
	}
	
	// ROW
	@Test
	public void testROW() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testROW(book);
	}

	// ROWS
	@Test
	public void testROWS() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testROWS(book);
	}

	// VLOOKUP
	@Test
	public void testVLOOKUP() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testVLOOKUP(book);
	}
}
