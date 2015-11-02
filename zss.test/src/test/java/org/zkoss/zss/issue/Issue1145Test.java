package org.zkoss.zss.issue;

import java.util.Collection;
import java.util.List;
import java.util.Locale;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;

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
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SBorder;
import org.zkoss.zss.model.SBorder.BorderType;
import org.zkoss.zss.model.SBorderLine;
import org.zkoss.zss.model.SCFValueObject;
import org.zkoss.zss.model.SCFValueObject.CFValueObjectType;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SColorScale;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.SConditionalFormattingRule.RuleOperator;
import org.zkoss.zss.model.SConditionalFormattingRule.RuleTimePeriod;
import org.zkoss.zss.model.SConditionalFormattingRule.RuleType;
import org.zkoss.zss.model.SDataBar;
import org.zkoss.zss.model.SExtraStyle;
import org.zkoss.zss.model.SFill;
import org.zkoss.zss.model.SFill.FillPattern;
import org.zkoss.zss.model.SFont.Boldweight;
import org.zkoss.zss.model.SFont.Underline;
import org.zkoss.zss.model.SIconSet.IconSetType;
import org.zkoss.zss.model.SIconSet;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SConditionalFormatting;
import org.zkoss.zss.model.impl.AbstractFillAdv;
import org.zkoss.zss.model.impl.AbstractFontAdv;
import org.zkoss.zss.model.impl.ColorImpl;

public class Issue1145Test {
	
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
	public void testImportConditionalFormatting() throws IOException {
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		ByteArrayInputStream is = null;
		try  {
			Book book0 = Util.loadBook(this, "book/1145-export-dxf.xlsx");
			Assert.assertTrue("No Exception", true);
			
			Exporters.getExporter("xlsx").export(book0, os);

			is = new ByteArrayInputStream(os.toByteArray());
			String bookName = book0.getBookName();
			Book book = Importers.getImporter().imports(is, bookName);
			
			List<SExtraStyle> styles = book.getInternalBook().getExtraStyles();
			Assert.assertEquals("number of SExtraStyles", 10, styles.size());
			
			test0(styles.get(0));
			test1_7_9(styles.get(1));
			test1_7_9(styles.get(2));
			test1_7_9(styles.get(3));
			test1_7_9(styles.get(4));
			test1_7_9(styles.get(5));
			test1_7_9(styles.get(6));
			test1_7_9(styles.get(7));
			test8(styles.get(8));
			test1_7_9(styles.get(9));
			
			File temp = Setup.getTempFile("Issue1145ExportDxf",".xlsx");
			Exporters.getExporter("xlsx").export(book0, temp);
		} catch (Exception e) {
			e.printStackTrace();
			Assert.assertTrue("Exception when load \"issue/book/1145-export-dxf.xlsx\":\n" + e, false);
		} finally {
			 os.close();
			 if (is!=null) is.close();
		}
	}

	//[0]
	private void test0(SExtraStyle s0) {
		final AbstractFontAdv font = (AbstractFontAdv) s0.getFont();
		final String numFmt = s0.getDataFormat();
		final AbstractFillAdv fill = (AbstractFillAdv) s0.getFill();
		final SBorder border = s0.getBorder();
		Assert.assertNotNull("Font", font);
		Assert.assertNotNull("numFmt", numFmt);
		Assert.assertNotNull("Fill", fill);
		Assert.assertNotNull("Border", border);
		
		//font
		Assert.assertTrue("bold", font.isOverrideBold());
		Assert.assertTrue("italic", font.isOverrideItalic());
		Assert.assertTrue("strike", font.isOverrideStrikeout());
		Assert.assertTrue("underline", font.isOverrideUnderline());
		Assert.assertTrue("color", font.isOverrideColor());
		Assert.assertFalse("heightPoint", font.isOverrideHeightPoints());
		Assert.assertFalse("name", font.isOverrideName());
		Assert.assertFalse("super/sub", font.isOverrideTypeOffset());
		
		Assert.assertEquals("getBold", Boldweight.BOLD, font.getBoldweight());
		Assert.assertEquals("isItalic", true, font.isItalic());
		Assert.assertEquals("isStrike",  true, font.isStrikeout());
		Assert.assertEquals("getUnderline", Underline.DOUBLE, font.getUnderline());
		Assert.assertEquals("getColor", new ColorImpl("00b050"), font.getColor());
		
		//numFmt
		Assert.assertEquals("numFmt", "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)", numFmt);
		
		//fill
		Assert.assertEquals("fillPatternType", FillPattern.LIGHT_DOWN, fill.getRawFillPattern());
		Assert.assertEquals("fillFgColor", new ColorImpl("ffff00"), fill.getRawFillColor());
		Assert.assertNull("fillBgColor", fill.getRawBackColor());
		
		//border
		final SBorderLine linel = border.getLeftLine();
		final SBorderLine liner = border.getRightLine();
		final SBorderLine linet = border.getTopLine();
		final SBorderLine lineb = border.getBottomLine();
		final SBorderLine lined = border.getDiagonalLine();
		final SBorderLine linev = border.getVerticalLine();
		final SBorderLine lineh = border.getHorizontalLine();
		
		Assert.assertNotNull("borderLeft", linel);
		Assert.assertNotNull("borderRight", liner);
		Assert.assertNotNull("borderTop", linet);
		Assert.assertNull("borderBottom", lineb);
		Assert.assertNull("borderDiagonal", lined);
		Assert.assertNull("borderVertical", linev);
		Assert.assertNull("borderHorizontal", lineh);
		
		Assert.assertEquals("borderLeftStyle", BorderType.THIN, linel.getBorderType());
		Assert.assertEquals("borderRightStyle", BorderType.THIN, liner.getBorderType());
		Assert.assertEquals("borderTopStyle", BorderType.THIN, linet.getBorderType());
		
		Assert.assertEquals("borderLeftColor", new ColorImpl("000000"), linel.getColor());
		Assert.assertEquals("borderRightStyle", new ColorImpl("000000"), liner.getColor());
		Assert.assertEquals("borderTopStyle", new ColorImpl("000000"), linet.getColor());
		
	}

	//[1~7, 9]
	private void test1_7_9(SExtraStyle s0) {
		final AbstractFontAdv font = (AbstractFontAdv) s0.getFont();
		final String numFmt = s0.getDataFormat();
		final AbstractFillAdv fill = (AbstractFillAdv) s0.getFill();
		final SBorder border = s0.getBorder();
		Assert.assertNotNull("Font", font);
		Assert.assertNull("numFmt", numFmt);
		Assert.assertNotNull("Fill", fill);
		Assert.assertNull("Border", border);
		
		//font
		Assert.assertFalse("bold", font.isOverrideBold());
		Assert.assertFalse("italic", font.isOverrideItalic());
		Assert.assertFalse("strike", font.isOverrideStrikeout());
		Assert.assertFalse("underline", font.isOverrideUnderline());
		Assert.assertTrue("color", font.isOverrideColor());
		Assert.assertFalse("heightPoint", font.isOverrideHeightPoints());
		Assert.assertFalse("name", font.isOverrideName());
		Assert.assertFalse("super/sub", font.isOverrideTypeOffset());
		
		Assert.assertEquals("getColor", new ColorImpl("9c0006"), font.getColor());
		
		//fill
		Assert.assertEquals("fillPatternType", FillPattern.NONE, fill.getRawFillPattern());
		Assert.assertNull("fillFgColor", fill.getRawFillColor());
		Assert.assertEquals("fillBgColor", new ColorImpl("ffc7ce"), fill.getRawBackColor());
	}

	//[8]
	private void test8(SExtraStyle s0) {
		final AbstractFontAdv font = (AbstractFontAdv) s0.getFont();
		final String numFmt = s0.getDataFormat();
		final AbstractFillAdv fill = (AbstractFillAdv) s0.getFill();
		final SBorder border = s0.getBorder();
		Assert.assertNotNull("Font", font);
		Assert.assertNull("numFmt", numFmt);
		Assert.assertNotNull("Fill", fill);
		Assert.assertNull("Border", border);
		
		//font
		Assert.assertFalse("bold", font.isOverrideBold());
		Assert.assertFalse("italic", font.isOverrideItalic());
		Assert.assertFalse("strike", font.isOverrideStrikeout());
		Assert.assertFalse("underline", font.isOverrideUnderline());
		Assert.assertTrue("color", font.isOverrideColor());
		Assert.assertFalse("heightPoint", font.isOverrideHeightPoints());
		Assert.assertFalse("name", font.isOverrideName());
		Assert.assertFalse("super/sub", font.isOverrideTypeOffset());
		
		Assert.assertEquals("getColor", new ColorImpl("9c6500"), font.getColor());
		
		//fill
		Assert.assertEquals("fillPatternType", FillPattern.NONE, fill.getRawFillPattern());
		Assert.assertNull("fillFgColor", fill.getRawFillColor());
		Assert.assertEquals("fillBgColor", new ColorImpl("ffeb9c"), fill.getRawBackColor());
	}
}
