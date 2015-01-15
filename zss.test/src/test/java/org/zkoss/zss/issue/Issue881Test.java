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
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.model.sys.format.FormatEngine;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.model.sys.format.FormatResult;
import org.zkoss.zss.model.impl.sys.FormatEngineImpl;

public class Issue881Test {
	
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
	 * Test a cell with text format should see =xxx a string rather than a 
	 * formula
	 */
	@Test
	public void testSpcialFormat() {
		FormatEngine formatEngine = new FormatEngineImpl();
		FormatContext ctx = new FormatContext(Setup.getZssLocale());
		FormatResult result = formatEngine.format("[DBNum1][$-404]General", 1d, ctx, 12);
		
		Assert.assertEquals("[DBNum1][$-404]General format", "1", result.getText());
		Assert.assertNotNull("JavaFormatter", result.getFormater());
	}

	@Test
	public void testGeneralFormat() {
		FormatEngine formatEngine = new FormatEngineImpl();
		FormatContext ctx = new FormatContext(Setup.getZssLocale());
		FormatResult result = formatEngine.format("General", 1d, ctx, 12);
		
		Assert.assertEquals("Genral", "1", result.getText());
		Assert.assertNotNull("JavaFormatter", result.getFormater());
	}
}
