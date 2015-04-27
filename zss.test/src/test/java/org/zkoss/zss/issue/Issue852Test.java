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

public class Issue852Test {
	
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
	 * parse a defined name which contains question mark '?'
	 * @throws IOException 
	 */
	@Test
	public void conditionalSumproduct() throws IOException {
		Book book = Util.loadBook(this, "book/852-conditional-sumproduct.xlsx");
		Sheet sheet1 = book.getSheetAt(0);

		Range rng = Ranges.range(sheet1, "I1");
		String text = rng.getCellFormatText();
		Assert.assertEquals("=SUMPRODUCT((F12:F21=1)*(G12:G21=\"Z\")*H12:H21)", "11", text);

		rng = Ranges.range(sheet1, "E1");
		text = rng.getCellFormatText();
		Assert.assertEquals("=SUMPRODUCT(F12:F21*H12:H21)", "106", text);

		rng = Ranges.range(sheet1, "D1");
		text = rng.getCellFormatText();
		Assert.assertEquals("=SUMPRODUCT(--(A1:A3=\"John\"),(B1:B3),(C1:C3))", "10", text);
		
		//+
		rng = Ranges.range(sheet1, "D2");
		text = rng.getCellFormatText();
		Assert.assertEquals("=B1:B3+C1:C3", "7", text);
		
		//-
		rng = Ranges.range(sheet1, "E2");
		text = rng.getCellFormatText();
		Assert.assertEquals("=B1:B3-C1:C3", "-3", text);

		//*
		rng = Ranges.range(sheet1, "F2");
		text = rng.getCellFormatText();
		Assert.assertEquals("=B1:B3*C1:C3", "10", text);

		//divide
		rng = Ranges.range(sheet1, "G2");
		text = rng.getCellFormatText();
		Assert.assertEquals("=B1:B3/C1:C3", "0.4", text);
		
		//power
		rng = Ranges.range(sheet1, "H2");
		text = rng.getCellFormatText();
		Assert.assertEquals("=power(B1:B3, C1:C3)", "32", text);
	}
}