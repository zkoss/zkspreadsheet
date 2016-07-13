package org.zkoss.zss.issue;

import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

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
import org.zkoss.zss.model.SIconSet.IconSetType;
import org.zkoss.zss.model.impl.AbstractBookAdv;
import org.zkoss.zss.model.SIconSet;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SConditionalFormatting;

public class Issue1141Test {
	
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
	public void testExportConditionalFormatting() throws IOException {
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		ByteArrayInputStream is;
		try  {
			Book book0 = Util.loadBook(this, "book/1141-export-conditional.xlsx");
			Assert.assertTrue("No Exception", true);
			
			Exporters.getExporter("xlsx").export(book0, os);

			is = new ByteArrayInputStream(os.toByteArray());
			String bookName = book0.getBookName();
			Book book = Importers.getImporter().imports(is, bookName);
			
			Sheet sheet1 = book.getSheetAt(0);
			SSheet sheet = sheet1.getInternalSheet();
			final List<SConditionalFormatting> list = sheet.getConditionalFormattings();
			SConditionalFormatting[] cfs = list.toArray(new SConditionalFormatting[list.size()]);
			Assert.assertEquals("number of <conditionalFormatting>", 20, cfs.length);
			
			AbstractBookAdv sbook = (AbstractBookAdv) book.getInternalBook();
			test0(sbook, cfs[0]);
			test1(sbook, cfs[1]);
			test2(sbook, cfs[2]);
			test3(sbook, cfs[3]);
			test4(sbook, cfs[4]);
			test5(sbook, cfs[5]);
			test6(sbook, cfs[6]);
			test7(sbook, cfs[7]);
			test8(sbook, cfs[8]);
			test9(sbook, cfs[9]);
			test10(sbook, cfs[10]);
			test11(sbook, cfs[11]);
			test12(sbook, cfs[12]);
			test13(sbook, cfs[13]);
			test14(sbook, cfs[14]);
			test15(sbook, cfs[15]);
			test16(sbook, cfs[16]);
			test17(sbook, cfs[17]);
			test18(sbook, cfs[18]);
			test19(sbook, cfs[19]);
			
		} catch (Exception e) {
			e.printStackTrace();
			Assert.assertTrue("Exception when load \"issue/book/1141-import-conditional.xlsx\":\n" + e, false);
		} finally {
			 os.close();
		}
	}

	private void testDxfId(AbstractBookAdv book, Long expect, SExtraStyle style) {
		if (expect == null) {
			Assert.assertNull(style);
		} else {
			Assert.assertNotNull(style);
			final int actual = book.indexOfExtraStyle(style);
			Assert.assertEquals("dxfId", expect.intValue(), actual);
		}
	}
	
	//[0]
	private void test0(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("A1:A10",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.CELL_IS, rule00.getType());
		Assert.assertEquals("rule's priority", 20, rule00.getPriority().intValue());
		Assert.assertEquals("rule's operator", RuleOperator.GREATER_THAN, rule00.getOperator());
		Assert.assertNotNull(rule00.getFormula1());
		Assert.assertNull(rule00.getFormula2());
		Assert.assertNull(rule00.getFormula3());
		final String formula00 = rule00.getFormula1();
		Assert.assertEquals("rule's formula", "=80", formula00);
		
		testDxfId(book, Long.valueOf(1), rule00.getExtraStyle());
	}

	//[1]
	private void test1(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("A8:A10",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.TIME_PERIOD, rule00.getType());
		Assert.assertEquals("rule's priority", 19, rule00.getPriority().intValue());
		Assert.assertEquals("rule's timePeriod", RuleTimePeriod.YESTERDAY, rule00.getTimePeriod());
		Assert.assertNotNull(rule00.getFormula1());
		Assert.assertNull(rule00.getFormula2());
		Assert.assertNull(rule00.getFormula3());
		final String formula00 = rule00.getFormula1();
		Assert.assertEquals("rule's formula", "=FLOOR(A8,1)=TODAY()-1", formula00);

		testDxfId(book, Long.valueOf(2), rule00.getExtraStyle());
	}

	//[2]
	private void test2(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("B1:B4",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.ICON_SET, rule00.getType());
		Assert.assertEquals("rule's priority", 18, rule00.getPriority().intValue());
		final SIconSet iconSet = rule00.getIconSet();
		Assert.assertEquals("iconSet type", IconSetType.X_3_TRAFFIC_LIGHTS_2, iconSet.getType());
		final List<SCFValueObject> vos = iconSet.getCFValueObjects();
		Assert.assertEquals("number of ValueObjects",  3, vos.size());
		SCFValueObject vo0 = vos.get(0);
		Assert.assertEquals("vo0 type",  CFValueObjectType.PERCENT, vo0.getType());
		Assert.assertEquals("vo0 value",  "0", vo0.getValue());
		SCFValueObject vo1 = vos.get(1);
		Assert.assertEquals("vo1 type",  CFValueObjectType.PERCENT, vo1.getType());
		Assert.assertEquals("vo1 value",  "33", vo1.getValue());
		SCFValueObject vo2 = vos.get(2);
		Assert.assertEquals("vo2 type",  CFValueObjectType.PERCENT, vo2.getType());
		Assert.assertEquals("vo2 value",  "67", vo2.getValue());
		
		testDxfId(book, null, rule00.getExtraStyle());

	}

	//[3]
	private void test3(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("C1:C6",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.TOP_10, rule00.getType());
		Assert.assertEquals("rule's priority", 17, rule00.getPriority().intValue());
		Assert.assertEquals("rule's rank", 10, rule00.getRank().longValue());
		
		testDxfId(book, Long.valueOf(1), rule00.getExtraStyle());
	}

	//[4]
	private void test4(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("D1:D4",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.COLOR_SCALE, rule00.getType());
		Assert.assertEquals("rule's priority", 16, rule00.getPriority().intValue());
		final SColorScale colorScale = rule00.getColorScale();
		final List<SCFValueObject> vos = colorScale.getCFValueObjects();
		Assert.assertEquals("number of ValueObjects",  3, vos.size());
		SCFValueObject vo0 = vos.get(0);
		Assert.assertEquals("vo0 type",  CFValueObjectType.MIN, vo0.getType());
		Assert.assertEquals("vo0 value",  "0", vo0.getValue());
		SCFValueObject vo1 = vos.get(1);
		Assert.assertEquals("vo1 type",  CFValueObjectType.PERCENTILE, vo1.getType());
		Assert.assertEquals("vo1 value",  "50", vo1.getValue());
		SCFValueObject vo2 = vos.get(2);
		Assert.assertEquals("vo2 type",  CFValueObjectType.MAX, vo2.getType());
		Assert.assertEquals("vo2 value",  "0", vo2.getValue());
		
		final List<SColor> colors = colorScale.getColors();
		Assert.assertEquals("number of Colors",  3, colors.size());
		SColor clr0 = colors.get(0);
		Assert.assertEquals("color0 code",  "#F8696B".toLowerCase(), clr0.getHtmlColor());
		
		SColor clr1 = colors.get(1);
		Assert.assertEquals("color1 code",  "#FFEB84".toLowerCase(), clr1.getHtmlColor());
		
		SColor clr2 = colors.get(2);
		Assert.assertEquals("color2 code",  "#63BE7B".toLowerCase(), clr2.getHtmlColor());
		
		testDxfId(book, null, rule00.getExtraStyle());
	}

	//[5]
	private void test5(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("E1:E9",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.DATA_BAR, rule00.getType());
		Assert.assertEquals("rule's priority", 15, rule00.getPriority().intValue());
		final SDataBar dataBar = rule00.getDataBar();
		final List<SCFValueObject> vos = dataBar.getCFValueObjects();
		Assert.assertEquals("number of ValueObjects",  2, vos.size());
		SCFValueObject vo0 = vos.get(0);
		Assert.assertEquals("vo0 type",  CFValueObjectType.MIN, vo0.getType());
		Assert.assertEquals("vo0 value",  "0", vo0.getValue());
		SCFValueObject vo1 = vos.get(1);
		Assert.assertEquals("vo1 type",  CFValueObjectType.MAX, vo1.getType());
		Assert.assertEquals("vo1 value",  "0", vo1.getValue());
		
		final SColor clr0 = dataBar.getColor();
		Assert.assertEquals("color0 code",  "#638ec6", clr0.getHtmlColor());
		
		testDxfId(book, null, rule00.getExtraStyle());
	}

	//[6]
	private void test6(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("F1:F6",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.TOP_10, rule00.getType());
		Assert.assertEquals("rule's priority", 14, rule00.getPriority().intValue());
		Assert.assertEquals("rule's rank", 10, rule00.getRank().longValue());
		Assert.assertEquals("rule's percent", true, rule00.isPercent());
		
		testDxfId(book, Long.valueOf(1), rule00.getExtraStyle());
	}

	//[7]
	private void test7(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("G1:G8",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.TOP_10, rule00.getType());
		Assert.assertEquals("rule's priority", 13, rule00.getPriority().intValue());
		Assert.assertEquals("rule's rank", 10, rule00.getRank().longValue());
		Assert.assertEquals("rule's bottom", true, rule00.isBottom());
		
		testDxfId(book, Long.valueOf(1), rule00.getExtraStyle());
	}

	//[8]
	private void test8(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("H1:H6",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.CONTAINS_TEXT, rule00.getType());
		Assert.assertEquals("rule's priority", 12, rule00.getPriority().intValue());
		Assert.assertEquals("rule's text", "Henri", rule00.getText());
		Assert.assertNotNull(rule00.getFormula1());
		Assert.assertNull(rule00.getFormula2());
		Assert.assertNull(rule00.getFormula3());
		final String formula00 = rule00.getFormula1();
		Assert.assertEquals("rule's formula", "=NOT(ISERROR(SEARCH(\"Henri\",H1)))", formula00);
		
		testDxfId(book, Long.valueOf(1), rule00.getExtraStyle());
	}

	//[9]
	private void test9(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("I1:I10",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.ABOVE_AVERAGE, rule00.getType());
		Assert.assertEquals("rule's priority", 11, rule00.getPriority().intValue());
		Assert.assertEquals("rule's aboveAverage", false, rule00.isAboveAverage());
		
		testDxfId(book, Long.valueOf(1), rule00.getExtraStyle());
	}
	
	//[10]
	private void test10(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("J3:J7",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.CELL_IS, rule00.getType());
		Assert.assertEquals("rule's priority", 10, rule00.getPriority().intValue());
		Assert.assertEquals("rule's operator", RuleOperator.BETWEEN, rule00.getOperator());
		Assert.assertNotNull(rule00.getFormula1());
		Assert.assertNotNull(rule00.getFormula2());
		Assert.assertNull(rule00.getFormula3());
		final String formula00 = rule00.getFormula1();
		Assert.assertEquals("rule's formula0", "=11", formula00);
		final String formula01 = rule00.getFormula2();
		Assert.assertEquals("rule's formula1", "=23", formula01);
		
		testDxfId(book, Long.valueOf(1), rule00.getExtraStyle());
	}
	
	//[11]
	private void test11(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("K3:K8",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.ABOVE_AVERAGE, rule00.getType());
		Assert.assertEquals("rule's priority", 9, rule00.getPriority().intValue());
		Assert.assertEquals("rule's equalAverage", true, rule00.isEqualAverage());
		
		testDxfId(book, null, rule00.getExtraStyle());
	}
	
	//[12]
	private void test12(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("L3:L7",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.ABOVE_AVERAGE, rule00.getType());
		Assert.assertEquals("rule's priority", 8, rule00.getPriority().intValue());
		Assert.assertEquals("rule's stdDev", 1, rule00.getStandardDeviation().intValue());
		
		testDxfId(book, null, rule00.getExtraStyle());
	}
	
	//[13]
	private void test13(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("M2:M7",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.DUPLICATE_VALUES, rule00.getType());
		Assert.assertEquals("rule's priority", 7, rule00.getPriority().intValue());
		
		testDxfId(book, Long.valueOf(1), rule00.getExtraStyle());
	}

	//[14]
	private void test14(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("N2:N11",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.ABOVE_AVERAGE, rule00.getType());
		Assert.assertEquals("rule's priority", 6, rule00.getPriority().intValue());
		Assert.assertEquals("rule's aboveAverage", false, rule00.isAboveAverage());
		Assert.assertEquals("rule's stdDev", 2, rule00.getStandardDeviation().intValue());
		
		testDxfId(book, null, rule00.getExtraStyle());
	}

	//[15]
	private void test15(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("A20:A26",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.EXPRESSION, rule00.getType());
		Assert.assertEquals("rule's priority", 5, rule00.getPriority().intValue());
		Assert.assertNotNull(rule00.getFormula1());
		Assert.assertNull(rule00.getFormula2());
		Assert.assertNull(rule00.getFormula3());
		final String formula00 = rule00.getFormula1();
		Assert.assertEquals("rule's formula", "=131", formula00);
		
		testDxfId(book, null, rule00.getExtraStyle());

	}

	//[16]
	private void test16(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("B20:B25",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.CONTAINS_BLANKS, rule00.getType());
		Assert.assertEquals("rule's priority", 4, rule00.getPriority().intValue());
		Assert.assertNotNull(rule00.getFormula1());
		Assert.assertNull(rule00.getFormula2());
		Assert.assertNull(rule00.getFormula3());
		final String formula00 = rule00.getFormula1();
		Assert.assertEquals("rule's formula", "=LEN(TRIM(B20))=0", formula00);
		
		testDxfId(book, null, rule00.getExtraStyle());
	}
	
	//[17]
	private void test17(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("C20:C24",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.CONTAINS_ERRORS, rule00.getType());
		Assert.assertEquals("rule's priority", 3, rule00.getPriority().intValue());
		Assert.assertNotNull(rule00.getFormula1());
		Assert.assertNull(rule00.getFormula2());
		Assert.assertNull(rule00.getFormula3());
		final String formula00 = rule00.getFormula1();
		Assert.assertEquals("rule's formula", "=ISERROR(C20)", formula00);

		testDxfId(book, null, rule00.getExtraStyle());
	}

	//[18]
	private void test18(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("D20:D24",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.NOT_CONTAINS_TEXT, rule00.getType());
		Assert.assertEquals("rule's priority", 2, rule00.getPriority().intValue());
		Assert.assertEquals("rule's operator", RuleOperator.NOT_CONTAINS, rule00.getOperator());
		Assert.assertEquals("rule's text", "Hello", rule00.getText());
		Assert.assertNotNull(rule00.getFormula1());
		Assert.assertNull(rule00.getFormula2());
		Assert.assertNull(rule00.getFormula3());
		final String formula00 = rule00.getFormula1();
		Assert.assertEquals("rule's formula", "=ISERROR(SEARCH(\"Hello\",D20))", formula00);
		
		testDxfId(book, Long.valueOf(0), rule00.getExtraStyle());
	}

	//[19]
	private void test19(AbstractBookAdv book, SConditionalFormatting cf0) {
		final Set<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.iterator().next();
		Assert.assertEquals("E20:E24",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.BEGINS_WITH, rule00.getType());
		Assert.assertEquals("rule's priority", 1, rule00.getPriority().intValue());
		Assert.assertEquals("rule's operator", RuleOperator.BEGINS_WITH, rule00.getOperator());
		Assert.assertEquals("rule's text", "xyz", rule00.getText());
		Assert.assertNotNull(rule00.getFormula1());
		Assert.assertNull(rule00.getFormula2());
		Assert.assertNull(rule00.getFormula3());
		final String formula00 = rule00.getFormula1();
		Assert.assertEquals("rule's formula", "=LEFT(E20,3)=\"xyz\"", formula00);
		
		testDxfId(book, null, rule00.getExtraStyle());
	}
}
