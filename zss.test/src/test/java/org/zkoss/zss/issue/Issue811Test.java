package org.zkoss.zss.issue;

import java.io.IOException;
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
import org.zkoss.zss.api.model.Sheet;

public class Issue811Test {
	
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

	//ref. http://www.currencysymbols.in/
	String[][] currencys = new String[][] {
			new String[] {"F1", "1"},
			new String[] {"F2", "2"},
			new String[] {"F3", "6"},
			new String[] {"F4", "6"},
			new String[] {"F5", "7"},
			new String[] {"F6", "8"},
			new String[] {"F7", "#N/A"},
			new String[] {"F8", "#N/A"},
			new String[] {"F9", "#N/A"},
	};
	
	/**
	 * parse a defined name which contains question mark '?'
	 * @throws IOException 
	 */
	@Test
	public void currencySymbols() throws IOException {
		Book book = Util.loadBook(this, "book/811-match-wildcard.xlsx");
		Sheet sheet1 = book.getSheetAt(0);
		for (String[] currency : currencys) {
			Range rng = Ranges.range(sheet1, currency[0]);
			String text = rng.getCellFormatText();
			Assert.assertEquals(currency[0], currency[1], text);
		}
	}
}