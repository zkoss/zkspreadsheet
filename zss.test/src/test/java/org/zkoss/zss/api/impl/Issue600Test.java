package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.UnsupportedEncodingException;
import java.util.Locale;

import org.junit.*;
import org.zkoss.poi.ss.usermodel.ErrorConstants;
import org.zkoss.zss.*;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.model.*;

/**
 * @author Hawk
 *
 */
public class Issue600Test {
	
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
	
	
	@Test
	public void testZSS610(){
		Book book;
		book = Util.loadBook(this,"book/blank.xlsx");
		Assert.assertEquals(Book.BookType.XLSX,book.getType());
		book = Util.loadBook(this,"book/blank.xls");
		Assert.assertEquals(Book.BookType.XLS,book.getType());
		
		
	}
	
	
}
