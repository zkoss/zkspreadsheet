package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Logical2007Test extends FormulaTestBase {
	
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
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testAND(book);
	}

	// FALSE
	@Test
	public void testFALSE() {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testFALSE(book);
	}

	// IF
	@Test
	public void testIF() {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testIF(book);
	}

	// IFERROR
	@Test
	public void testIFERROR() {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testIFERROR(book);
	}

	// NOT
	@Test
	public void testNOT() {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testNOT(book);
	}

	// OR
	@Test
	public void testOR() {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testOR(book);
	}

	// TRUE
	@Test
	public void testTRUE() {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testTRUE(book);
	}
}
