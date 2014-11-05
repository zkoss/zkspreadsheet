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

public class Issue777Test {
	
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
			new String[] {"B1", "NT$100.00"},
			new String[] {"B2", "$200.00"},
			new String[] {"B3", "¥300.00"},
			new String[] {"B4", "¥100.00"},
			new String[] {"B5", "₩100.00"},
			new String[] {"B6", "€ 100.00"},
			new String[] {"B7", "HKD 100.00"},
			new String[] {"B8", "RM100.00"},
			new String[] {"B9", "HK$100.00"},
			new String[] {"B10", "NT$100.00"},
	};
	
	/**
	 * parse a defined name which contains question mark '?'
	 * @throws IOException 
	 */
	@Test
	public void currencySymbols() throws IOException {
		Book book = Util.loadBook(this, "book/777-currencySymbol.xlsx");
		Sheet sheet1 = book.getSheetAt(0);
		for (String[] currency : currencys) {
			Range rng = Ranges.range(sheet1, currency[0]);
			String text = rng.getCellFormatText();
			Assert.assertEquals(currency[0], currency[1], text);
		}
	}
}