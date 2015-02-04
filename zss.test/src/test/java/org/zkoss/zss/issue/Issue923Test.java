package org.zkoss.zss.issue;

import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.model.SBook;

public class Issue923Test {
	
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
	 * Test dirty checking in book
	 */
	@Test
	public void testDirty() {
		Book book = Util.loadBook(this, "book/blank.xlsx");
		SBook sbook = book.getInternalBook();
		Assert.assertFalse("Dirty should be false right after book loaded from importer.", sbook.isDirty());
		
		Range range = Ranges.range(book.getSheetAt(0));
		range.cloneSheet("sheet-test");
		Assert.assertTrue("Dirty should be true after modify via range.", sbook.isDirty());
		
		sbook.setDirty(false);
		Assert.assertFalse("Dirty should be false after reset.", sbook.isDirty());
		
		range.clearAll();
		Assert.assertTrue("Dirty should be true after modify via range.", sbook.isDirty());
	}
}
