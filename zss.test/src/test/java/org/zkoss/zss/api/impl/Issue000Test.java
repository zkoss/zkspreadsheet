package org.zkoss.zss.api.impl;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
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
		ZssContext.setThreadLocal(new ZssContext(Locale.TAIWAN,-1));
	}
	
	@After
	public void tearDown() throws Exception {
		ZssContext.setThreadLocal(null);
	}
	
	/**
	 * Exception when exporting excel twice.
	 */
	@Test
	public void testZSS36() throws IOException {
		final String filename = "book/blank.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Util.export(workbook, Setup.getTempFile());
		Util.export(workbook, Setup.getTempFile());
	}
	
}
