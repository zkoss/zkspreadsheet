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

public class Issue847Test {
	
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

	//ref. http://office.microsoft.com/zh-tw/excel-help/HA102752975.aspx
	String[][] values = new String[][] {
			new String[] {"A1", "1"},
			new String[] {"A2", "2"},
			new String[] {"B1", "3"},
	};
	
	/**
	 * parse a defined name which contains question mark '?'
	 * @throws IOException 
	 */
	@Test
	public void parenthesisAdd() throws IOException {
		Book book = Util.loadBook(this, "book/847-parenthesis-add.xlsx");
		Sheet sheet1 = book.getSheetAt(0);
		for (String[] currency : values) {
			Range rng = Ranges.range(sheet1, currency[0]);
			String text = rng.getCellFormatText();
			Assert.assertEquals(currency[0], currency[1], text);
		}
		Range rngB1 = Ranges.range(sheet1, "B1");
		Assert.assertEquals("formula", "=(A1+A2)", rngB1.getCellEditText());
	}
}