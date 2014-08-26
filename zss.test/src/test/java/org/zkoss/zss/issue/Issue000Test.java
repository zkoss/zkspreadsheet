package org.zkoss.zss.issue;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

/**
 * ZSS-36
 * @author kuro
 */
public class Issue000Test {

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
	
	/**
	 * Exception when exporting excel twice.
	 */
	@Test 
	public void testZSS36() throws IOException {
		Book workbook = Util.loadBook(this, "book/blank.xlsx");
		Util.export(workbook, Setup.getTempFile());
		Util.export(workbook, Setup.getTempFile());
	}
	
}
