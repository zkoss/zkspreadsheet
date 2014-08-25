package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Logical2003Test extends FormulaTestBase {
	
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
	
	// AND
	@Test
	public void testAND() {
		Book book = Util.loadBook(this, "book/TestFile2003-Formula.xls");
		testAND(book);
	}

	// FALSE
	@Test
	public void testFALSE() {
		Book book = Util.loadBook(this, "book/TestFile2003-Formula.xls");
		testFALSE(book);
	}

	// IF
	@Test
	public void testIF() {
		Book book = Util.loadBook(this, "book/TestFile2003-Formula.xls");
		testIF(book);
	}

	// IFERROR
	@Test
	public void testIFERROR() {
		Book book = Util.loadBook(this, "book/TestFile2003-Formula.xls");
		testIFERROR(book);
	}

	// NOT
	@Test
	public void testNOT() {
		Book book = Util.loadBook(this, "book/TestFile2003-Formula.xls");
		testNOT(book);
	}

	// OR
	@Test
	public void testOR() {
		Book book = Util.loadBook(this, "book/TestFile2003-Formula.xls");
		testOR(book);
	}

	// TRUE
	@Test
	public void testTRUE() {
		Book book = Util.loadBook(this, "book/TestFile2003-Formula.xls");
		testTRUE(book);
	}
}
