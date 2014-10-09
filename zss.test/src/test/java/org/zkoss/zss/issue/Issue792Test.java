package org.zkoss.zss.issue;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.impl.pdf.PdfExporter;

public class Issue792Test {
	
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
	 * Test render long text that across the cell boundary
	 */
	@Test
	public void testCheckLegalName() {
		try {
			Book book = Util.loadBook(this, "book/790-legal-name-name.xlsx");
			Assert.assertTrue(true);
		} catch (Exception ex) {
			Assert.assertTrue("Check Name name:" + ex.getMessage(), false);
		}
	}
}
