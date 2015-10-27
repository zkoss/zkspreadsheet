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
import org.zkoss.zss.model.SIconSet.IconSetType;
import org.zkoss.zss.model.SIconSet;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SConditionalFormatting;

public class Issue1130Test {
	
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
		try  {
			Book book = Util.loadBook(this, "book/1130-import-conditional.xlsx");
			Assert.assertTrue("No Exception", true);
			Sheet sheet1 = book.getSheetAt(0);
			SSheet sheet = sheet1.getInternalSheet();
			final List<SConditionalFormatting> list = sheet.getConditonalFormattings();
			SConditionalFormatting[] cfs = list.toArray(new SConditionalFormatting[list.size()]);
			Assert.assertEquals("number of <conditionalFormatting>", 20, cfs.length);
			
			test0(cfs[0]);
			test1(cfs[1]);
			test2(cfs[2]);
			test3(cfs[3]);
			test4(cfs[4]);
			test5(cfs[5]);
			test6(cfs[6]);
			test7(cfs[7]);
			test8(cfs[8]);
			test9(cfs[9]);
			test10(cfs[10]);
			test11(cfs[11]);
			test12(cfs[12]);
			test13(cfs[13]);
			test14(cfs[14]);
			test15(cfs[15]);
			test16(cfs[16]);
			test17(cfs[17]);
			test18(cfs[18]);
			test19(cfs[19]);
			
		} catch (Exception e) {
			e.printStackTrace();
			Assert.assertTrue("Exception when load \"issue/book/1141-import-conditional.xlsx\":\n" + e, false);
		} finally {
			 os.close();
		}
	}

	//[0]
	private void test0(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("A1:A10",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.CELL_IS, rule00.getType());
		Assert.assertEquals("rule's priority", 20, rule00.getPriority());
		Assert.assertEquals("rule's operator", RuleOperator.GREATER_THAN, rule00.getOperator());
		final List<String> formulas = rule00.getFormulas();
		Assert.assertEquals("number of formulas",  1, formulas.size());
		final String formula00 = formulas.get(0);
		Assert.assertEquals("rule's formula", "80", formula00);
	}

	//[1]
	private void test1(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("A8:A10",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.TIME_PERIOD, rule00.getType());
		Assert.assertEquals("rule's priority", 19, rule00.getPriority());
		Assert.assertEquals("rule's timePeriod", RuleTimePeriod.YESTERDAY, rule00.getTimePeriod());
		final List<String> formulas = rule00.getFormulas();
		Assert.assertEquals("number of formulas",  1, formulas.size());
		final String formula00 = formulas.get(0);
		Assert.assertEquals("rule's formula", "FLOOR(A8,1)=TODAY()-1", formula00);
	}

	//[2]
	private void test2(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("B1:B4",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.ICON_SET, rule00.getType());
		Assert.assertEquals("rule's priority", 18, rule00.getPriority());
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
	}

	//[3]
	private void test3(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("C1:C6",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.TOP_10, rule00.getType());
		Assert.assertEquals("rule's priority", 17, rule00.getPriority());
		Assert.assertEquals("rule's rank", 10, rule00.getRank());
	}

	//[4]
	private void test4(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("D1:D4",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.COLOR_SCALE, rule00.getType());
		Assert.assertEquals("rule's priority", 16, rule00.getPriority());
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
	}

	//[5]
	private void test5(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("E1:E9",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.DATA_BAR, rule00.getType());
		Assert.assertEquals("rule's priority", 15, rule00.getPriority());
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
	}

	//[6]
	private void test6(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("F1:F6",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.TOP_10, rule00.getType());
		Assert.assertEquals("rule's priority", 14, rule00.getPriority());
		Assert.assertEquals("rule's rank", 10, rule00.getRank());
		Assert.assertEquals("rule's percent", true, rule00.isPercent());
	}

	//[7]
	private void test7(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("G1:G8",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.TOP_10, rule00.getType());
		Assert.assertEquals("rule's priority", 13, rule00.getPriority());
		Assert.assertEquals("rule's rank", 10, rule00.getRank());
		Assert.assertEquals("rule's bottom", true, rule00.isBottom());
	}

	//[8]
	private void test8(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("H1:H6",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.CONTAINS_TEXT, rule00.getType());
		Assert.assertEquals("rule's priority", 12, rule00.getPriority());
		Assert.assertEquals("rule's text", "Henri", rule00.getText());
		final List<String> formulas = rule00.getFormulas();
		Assert.assertEquals("number of formulas",  1, formulas.size());
		final String formula00 = formulas.get(0);
		Assert.assertEquals("rule's formula", "NOT(ISERROR(SEARCH(\"Henri\",H1)))", formula00);
	}

	//[9]
	private void test9(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("I1:I10",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.ABOVE_AVERAGE, rule00.getType());
		Assert.assertEquals("rule's priority", 11, rule00.getPriority());
		Assert.assertEquals("rule's aboveAverage", false, rule00.isAboveAverage());
	}
	
	//[10]
	private void test10(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("J3:J7",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.CELL_IS, rule00.getType());
		Assert.assertEquals("rule's priority", 10, rule00.getPriority());
		Assert.assertEquals("rule's operator", RuleOperator.BETWEEN, rule00.getOperator());
		final List<String> formulas = rule00.getFormulas();
		Assert.assertEquals("number of formulas",  2, formulas.size());
		final String formula00 = formulas.get(0);
		Assert.assertEquals("rule's formula0", "11", formula00);
		final String formula01 = formulas.get(1);
		Assert.assertEquals("rule's formula1", "23", formula01);
	}
	
	//[11]
	private void test11(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("K3:K8",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.ABOVE_AVERAGE, rule00.getType());
		Assert.assertEquals("rule's priority", 9, rule00.getPriority());
		Assert.assertEquals("rule's equalAverage", true, rule00.isEqualAverage());
	}
	
	//[12]
	private void test12(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("L3:L7",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.ABOVE_AVERAGE, rule00.getType());
		Assert.assertEquals("rule's priority", 8, rule00.getPriority());
		Assert.assertEquals("rule's stdDev", 1, rule00.getStandardDeviation());
	}
	
	//[13]
	private void test13(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("M2:M7",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.DUPLICATE_VALUES, rule00.getType());
		Assert.assertEquals("rule's priority", 7, rule00.getPriority());
	}

	//[14]
	private void test14(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("N2:N11",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.ABOVE_AVERAGE, rule00.getType());
		Assert.assertEquals("rule's priority", 6, rule00.getPriority());
		Assert.assertEquals("rule's aboveAverage", false, rule00.isAboveAverage());
		Assert.assertEquals("rule's stdDev", 2, rule00.getStandardDeviation());
	}

	//[15]
	private void test15(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("A20:A26",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.EXPRESSION, rule00.getType());
		Assert.assertEquals("rule's priority", 5, rule00.getPriority());
		final List<String> formulas = rule00.getFormulas();
		Assert.assertEquals("number of formulas",  1, formulas.size());
		final String formula00 = formulas.get(0);
		Assert.assertEquals("rule's formula", "131", formula00);
	}

	//[16]
	private void test16(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("B20:B25",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.CONTAINS_BLANKS, rule00.getType());
		Assert.assertEquals("rule's priority", 4, rule00.getPriority());
		final List<String> formulas = rule00.getFormulas();
		Assert.assertEquals("number of formulas",  1, formulas.size());
		final String formula00 = formulas.get(0);
		Assert.assertEquals("rule's formula", "LEN(TRIM(B20))=0", formula00);
	}
	
	//[17]
	private void test17(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("C20:C24",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.CONTAINS_ERRORS, rule00.getType());
		Assert.assertEquals("rule's priority", 3, rule00.getPriority());
		final List<String> formulas = rule00.getFormulas();
		Assert.assertEquals("number of formulas",  1, formulas.size());
		final String formula00 = formulas.get(0);
		Assert.assertEquals("rule's formula", "ISERROR(C20)", formula00);
	}

	//[18]
	private void test18(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("D20:D24",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.NOT_CONTAINS_TEXT, rule00.getType());
		Assert.assertEquals("rule's priority", 2, rule00.getPriority());
		Assert.assertEquals("rule's operator", RuleOperator.NOT_CONTAINS, rule00.getOperator());
		Assert.assertEquals("rule's text", "Hello", rule00.getText());
		final List<String> formulas = rule00.getFormulas();
		Assert.assertEquals("number of formulas",  1, formulas.size());
		final String formula00 = formulas.get(0);
		Assert.assertEquals("rule's formula", "ISERROR(SEARCH(\"Hello\",D20))", formula00);
	}

	//[19]
	private void test19(SConditionalFormatting cf0) {
		final List<CellRegion> regions = cf0.getRegions();
		Assert.assertEquals("number of regions",  1, regions.size());
		final CellRegion rgn00 = regions.get(0);
		Assert.assertEquals("E20:E24",  rgn00.getReferenceString());
		final List<SConditionalFormattingRule> rules = cf0.getRules();
		Assert.assertEquals("number of rules",  1, rules.size());
		final SConditionalFormattingRule rule00 = rules.get(0);
		Assert.assertEquals("rule's type", RuleType.BEGINS_WITH, rule00.getType());
		Assert.assertEquals("rule's priority", 1, rule00.getPriority());
		Assert.assertEquals("rule's operator", RuleOperator.BEGINS_WITH, rule00.getOperator());
		Assert.assertEquals("rule's text", "xyz", rule00.getText());
		final List<String> formulas = rule00.getFormulas();
		Assert.assertEquals("number of formulas",  1, formulas.size());
		final String formula00 = formulas.get(0);
		Assert.assertEquals("rule's formula", "LEFT(E20,3)=\"xyz\"", formula00);
	}
}
