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
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.api.Exporters;

public class Issue1317Test {
	
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
		Book book = Util.loadBook(this, "book/1317-conditional-format.xlsx");
		Sheet sheet1 = book.getSheetAt(0);
		Range rng1 = Ranges.range(sheet1, "D5:J8");
		Range rng2 = Ranges.range(sheet1, "D10");
		Range rng3 = Ranges.range(sheet1, "D15");
		Range rng4 = Ranges.range(sheet1, "D20");
		rng1.paste(rng2);
		rng1.paste(rng3);
		rng1.paste(rng4);
		for (int chunk = 1; chunk <= 4; ++chunk) {
			for (int row0 = 0; row0 < 4; ++row0) {
				_equals(sheet1, chunk * 5 + row0);
			}
		}
	}
	void _equals(Sheet sheet1, int row) {
		SConditionalFormatting scf1 = ((SheetImpl)sheet1.getInternalSheet()).getConditionalFormatting(row - 1, 5); //F5
		List<SConditionalFormattingRule> scfRules1 = scf1.getRules();
		Assert.assertEquals("F"+row+" formula1", "=F"+row+"-$E"+row+">5", scfRules1.get(0).getFormula1());
	}
}
 