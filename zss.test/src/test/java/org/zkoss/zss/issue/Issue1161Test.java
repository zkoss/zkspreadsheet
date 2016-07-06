package org.zkoss.zss.issue;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
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
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SConditionalFormatting;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.SConditionalFormattingRule.RuleType;
import org.zkoss.zss.model.SDataBar;
import org.zkoss.zss.model.SExtraStyle;
import org.zkoss.zss.model.SFill.FillPattern;
import org.zkoss.zss.model.SIconSet;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.impl.AbstractFillAdv;

public class Issue1161Test {
	
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
	 * Test a book with print area and export it and it should not throw exception
	 * @throws IOException 
	 */
	@Test
	public void testExportXlsBookWithPrintArea() throws IOException {
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		ByteArrayInputStream is;
		try  {
			Book book0 = Util.loadBook(this, "book/1161-export-showicon.xlsx");
			Assert.assertTrue("No Exception", true);
			Exporters.getExporter("xlsx").export(book0, os);
			
			is = new ByteArrayInputStream(os.toByteArray());
			String bookName = book0.getBookName();
			Book book = Importers.getImporter().imports(is, bookName);
			SSheet sheet = book.getSheetAt(0).getInternalSheet();
			
			final List<SConditionalFormatting> list = sheet.getConditionalFormattings();
			Assert.assertEquals("number of condition formatting", 6, list.size());
			for (SConditionalFormatting cf0 : list) {
				final List<SConditionalFormattingRule> rules = cf0.getRules();
				Assert.assertEquals("number of condition formatting rules",  1, rules.size());
				final SConditionalFormattingRule rule00 = rules.get(0);
				RuleType rtype =  rule00.getType();
				if (rtype == RuleType.DATA_BAR) {
					Assert.assertTrue("!DataBar#showIconOnly", rule00.getDataBar().isShowValue());
				} else if (rtype == RuleType.ICON_SET) {
					Assert.assertTrue("!IconBar#showIconOnly", rule00.getIconSet().isShowValue());
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			Assert.assertTrue("Exception when load \"issue/book/1161-export-showicon.xlsx\":\n" + e, false);
		} finally {
			 os.close();
		}
	}
}
