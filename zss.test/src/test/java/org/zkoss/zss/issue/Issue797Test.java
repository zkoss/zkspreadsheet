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
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.api.model.Sheet;

public class Issue797Test {
	
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
	 * Test export white background to black background
	 * @throws IOException 
	 */
	@Test
	public void testExportWhiteToBlackBackground() throws IOException {
		Book book = Util.loadBook(this, "book/797-export-black-background.xlsx");
		Sheet sheet1 = book.getSheet("Input");
		Range rngB1 = Ranges.range(sheet1, "B1");
		CellStyle style = rngB1.getCellStyle();
		Color fillColor = style.getFillColor();
		Color backColor = style.getBackColor();
		Assert.assertEquals("fill forecolor", "#ffffff", fillColor.getHtmlColor());
		Assert.assertEquals("fill backcolor", "#000000", backColor.getHtmlColor());
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet1 = book.getSheetAt(0);
		rngB1 = Ranges.range(sheet1, "B1");
		style = rngB1.getCellStyle();
		fillColor = style.getFillColor();
		backColor = style.getBackColor();
		Assert.assertEquals("fill forecolor", "#ffffff", fillColor.getHtmlColor());
		Assert.assertEquals("fill backcolor", "#000000", backColor.getHtmlColor());
	}
	
	@Test
	public void testImportBlackBackground() {
		Book book = Util.loadBook(this, "book/797-black-background.xlsx");
		Sheet sheet1 = book.getSheet("Input");
		Range rngB1 = Ranges.range(sheet1, "B1");
		CellStyle style = rngB1.getCellStyle();
		Color fillColor = style.getFillColor();
		Color backColor = style.getBackColor();
		Assert.assertEquals("fill forecolor", "#000000", fillColor.getHtmlColor());
		Assert.assertEquals("fill backcolor", "#000000", backColor.getHtmlColor());
	}
}
