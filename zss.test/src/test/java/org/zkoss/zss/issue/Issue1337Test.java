package org.zkoss.zss.issue;

import java.util.List;
import java.util.Locale;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.Validation;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SConditionalFormatting;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.impl.SheetImpl;
import org.zkoss.zss.model.SName;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.api.Exporters;

public class Issue1337Test {
	
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
	 * Test a book with conditinal format and copy paste
	 * @throws IOException 
	 */
	@Test
	public void testExportXlsBookWithPrintArea() throws IOException {
		Book book = Util.loadBook(this, "book/1337-extend-name.xlsx");
		Sheet sheet2 = book.getSheet("Sheet 2");
		SName name1 = book.getInternalBook().getNameByName("TestName", null); // workbook scope
		SName name2 = book.getInternalBook().getNameByName("TestName", "Sheet 2"); // sheet2 scope
		Assert.assertEquals("workbook scope TestName", "'Sheet 2'!$C$10:$H$10", name1.getRefersToFormula());
		Assert.assertEquals("worksheet scope TestName", "'Sheet 2'!$C$10:$H$10", name2.getRefersToFormula());
		Range rng1 = Ranges.range(sheet2, "A9").toRowRange();
		rng1.insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_NONE);
		SName name21 = book.getInternalBook().getNameByName("TestName", null); // workbook scope
		SName name22 = book.getInternalBook().getNameByName("TestName", "Sheet 2"); // sheet2 scope
		Assert.assertEquals("workbook scope TestName", "'Sheet 2'!$C$11:$H$11", name21.getRefersToFormula());
		Assert.assertEquals("worksheet scope TestName", "'Sheet 2'!$C$11:$H$11", name22.getRefersToFormula());
	}
}
 