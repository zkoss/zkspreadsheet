package org.zkoss.zss.api.impl;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.AssertUtil;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * Complex tests for issue 434
 * @author dennis
 *
 */
public class Issue434Test {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		ZssContext.setThreadLocal(new ZssContext(Locale.TAIWAN,-1));
	}
	
	@After
	public void tearDown() throws Exception {
		ZssContext.setThreadLocal(null);
	}
	
	@Test
	public void testBlank_InsertRow() throws IOException {
		testBlank_InsertRow1(Util.loadBook(this,"book/blank.xlsx"));
		testBlank_InsertRow2(Util.loadBook(this,"book/blank.xlsx"));
		testBlank_InsertRow1(Util.loadBook(this,"book/blank.xls"));
		testBlank_InsertRow2(Util.loadBook(this,"book/blank.xls"));
	}
	private void testBlank_InsertRow1(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
		
		Ranges.range(sheet,"B2:D4").merge(false);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Ranges.range(sheet,"A3").toRowRange().insert(Range.InsertShift.DOWN,Range.InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D5", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D5", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
	}
	private void testBlank_InsertRow2(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
		
		
		Ranges.range(sheet,"B2:D4").merge(false);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Ranges.range(sheet,"A2").toRowRange().insert(Range.InsertShift.DOWN,Range.InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B3:D5", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B3:D5", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
	}
	
	
	@Test
	public void testCase1_InsertRow() throws IOException {
		testCase1_InsertRow1(Util.loadBook(this,"book/434-case1.xlsx"));
		testCase1_InsertRow1(mergeForCase1(Util.loadBook(this,"book/blank.xlsx")));
		testCase1_InsertRow1(Util.loadBook(this,"book/434-case1.xls"));
		testCase1_InsertRow1(mergeForCase1(Util.loadBook(this,"book/blank.xls")));
	}
	private void testCase1_InsertRow1(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(4, sheet.getPoiSheet().getNumMergedRegions());
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);
		
		
		Ranges.range(sheet,"A6:A7").toRowRange().insert(Range.InsertShift.DOWN,Range.InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B12:D14","K12:M14"}, sheet);
		
		Ranges.range(sheet,"A12").toRowRange().insert(Range.InsertShift.DOWN,Range.InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B13:D15","K13:M15"}, sheet);
		
		Ranges.range(sheet,"A14").toRowRange().insert(Range.InsertShift.DOWN,Range.InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B13:D16","K13:M16"}, sheet);

	
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B13:D16","K13:M16"}, sheet);
	}
	
	
	@Test
	public void testBlank_DeleteRow() throws IOException {
		testBlank_DeleteRow1(Util.loadBook(this,"book/blank.xlsx"));
		testBlank_DeleteRow2(Util.loadBook(this,"book/blank.xlsx"));
		testBlank_DeleteRow1(Util.loadBook(this,"book/blank.xls"));
		testBlank_DeleteRow2(Util.loadBook(this,"book/blank.xls"));
	}
	private void testBlank_DeleteRow1(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
		
		Ranges.range(sheet,"B2:D4").merge(false);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Ranges.range(sheet,"A3").toRowRange().delete(DeleteShift.UP);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D3", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D3", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
	}
	private void testBlank_DeleteRow2(Book book) throws IOException {
		
		
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
		
		
		Ranges.range(sheet,"B2:D4").merge(false);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Ranges.range(sheet,"A2").toRowRange().delete(DeleteShift.UP);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D3", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D3", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
	}
	

	@Test
	public void testCase1_DeleteRow() throws IOException {
		testCase1_DeleteRow1(Util.loadBook(this,"book/434-case1.xlsx"));
		testCase1_DeleteRow1(mergeForCase1(Util.loadBook(this,"book/blank.xlsx")));
		testCase1_DeleteRow1(Util.loadBook(this,"book/434-case1.xls"));
		testCase1_DeleteRow1(mergeForCase1(Util.loadBook(this,"book/blank.xls")));
	}
	private void testCase1_DeleteRow1(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(4, sheet.getPoiSheet().getNumMergedRegions());
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);
		
		
		Ranges.range(sheet,"A6:A7").toRowRange().delete(Range.DeleteShift.UP);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B8:D10","K8:M10"}, sheet);
		
		Ranges.range(sheet,"A9").toRowRange().delete(Range.DeleteShift.UP);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B8:D9","K8:M9"}, sheet);
		
		Ranges.range(sheet,"A8").toRowRange().delete(Range.DeleteShift.UP);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B8:D8","K8:M8"}, sheet);
	
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B8:D8","K8:M8"}, sheet);
	}
	
	
	@Test
	public void testBlank_InserColumn() throws IOException {
		testBlank_InserColumn1(Util.loadBook(this,"book/blank.xlsx"));
		testBlank_InserColumn2(Util.loadBook(this,"book/blank.xlsx"));
		testBlank_InserColumn1(Util.loadBook(this,"book/blank.xls"));
		testBlank_InserColumn2(Util.loadBook(this,"book/blank.xls"));
	}
	private void testBlank_InserColumn1(Book book) throws IOException {
		
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
		
		Ranges.range(sheet,"B2:D4").merge(false);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Ranges.range(sheet,"C1").toColumnRange().insert(Range.InsertShift.RIGHT,Range.InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:E4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:E4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
	}
	private void testBlank_InserColumn2(Book book) throws IOException {
		
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
		
		Ranges.range(sheet,"B2:D4").merge(false);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Ranges.range(sheet,"B1").toColumnRange().insert(Range.InsertShift.RIGHT,Range.InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("C2:E4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("C2:E4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
	}
	
	
	@Test
	public void testBlank_DeleteColumn() throws IOException {
		testBlank_DeleteColumn1(Util.loadBook(this,"book/blank.xlsx"));
		testBlank_DeleteColumn2(Util.loadBook(this,"book/blank.xlsx"));
		testBlank_DeleteColumn1(Util.loadBook(this,"book/blank.xls"));
		testBlank_DeleteColumn2(Util.loadBook(this,"book/blank.xls"));
	}
	private void testBlank_DeleteColumn1(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
		
		Ranges.range(sheet,"B2:D4").merge(false);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Ranges.range(sheet,"C3").toColumnRange().delete(DeleteShift.LEFT);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:C4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:C4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
	}
	private void testBlank_DeleteColumn2(Book book) throws IOException {
		
		
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
		
		
		Ranges.range(sheet,"B2:D4").merge(false);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:D4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Ranges.range(sheet,"B2").toColumnRange().delete(DeleteShift.LEFT);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:C4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("B2:C4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
	}
	
	
	@Test
	public void testCase1_InsertColumn() throws IOException {
		testCase1_InsertColumn1(Util.loadBook(this,"book/434-case1.xlsx"));
		testCase1_InsertColumn1(mergeForCase1(Util.loadBook(this,"book/blank.xlsx")));
		testCase1_InsertColumn1(Util.loadBook(this,"book/434-case1.xls"));
		testCase1_InsertColumn1(mergeForCase1(Util.loadBook(this,"book/blank.xls")));
	}
	private void testCase1_InsertColumn1(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(4, sheet.getPoiSheet().getNumMergedRegions());
		
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);
		
		Ranges.range(sheet,"F1:G1").toColumnRange().insert(Range.InsertShift.RIGHT,Range.InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","M2:O4","B10:D12","M10:O12"}, sheet);
		
		Ranges.range(sheet,"M1").toColumnRange().insert(Range.InsertShift.RIGHT,Range.InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","N2:P4","B10:D12","N10:P12"}, sheet);
		
		Ranges.range(sheet,"O1").toColumnRange().insert(Range.InsertShift.RIGHT,Range.InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","N2:Q4","B10:D12","N10:Q12"}, sheet);
	
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","N2:Q4","B10:D12","N10:Q12"}, sheet);
	}
	@Test
	public void testCase1_DeleteColumn() throws IOException {
		testCase1_DeleteColumn1(Util.loadBook(this,"book/434-case1.xlsx"));
		testCase1_DeleteColumn1(mergeForCase1(Util.loadBook(this,"book/blank.xlsx")));
		testCase1_DeleteColumn1(Util.loadBook(this,"book/434-case1.xls"));
		testCase1_DeleteColumn1(mergeForCase1(Util.loadBook(this,"book/blank.xls")));
	}
	
	private Book mergeForCase1(Book book){
		Sheet sheet = book.getSheetAt(0);
		Ranges.range(sheet,"B2:D4").merge(false);
		Ranges.range(sheet,"K2:M4").merge(false);
		Ranges.range(sheet,"B10:D12").merge(false);
		Ranges.range(sheet,"K10:M12").merge(false);
		return book;
	}
	
	private void testCase1_DeleteColumn1(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(4, sheet.getPoiSheet().getNumMergedRegions());
		
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);
		
		Ranges.range(sheet,"F1:G1").toColumnRange().delete(Range.DeleteShift.LEFT);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","I2:K4","B10:D12","I10:K12"}, sheet);
		
		Ranges.range(sheet,"J1").toColumnRange().delete(Range.DeleteShift.LEFT);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","I2:J4","B10:D12","I10:J12"}, sheet);
		
		Ranges.range(sheet,"I1").toColumnRange().delete(Range.DeleteShift.LEFT);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","I2:I4","B10:D12","I10:I12"}, sheet);		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","I2:I4","B10:D12","I10:I12"}, sheet);
	}
	
	
	@Test
	public void testCase1_InsertRange() throws IOException {
		testCase1_InsertRange1(Util.loadBook(this,"book/434-case1.xlsx"));
		testCase1_InsertRange1(mergeForCase1(Util.loadBook(this,"book/blank.xlsx")));
		testCase1_InsertRange1(Util.loadBook(this,"book/434-case1.xls"));
		testCase1_InsertRange1(mergeForCase1(Util.loadBook(this,"book/blank.xls")));
	}
	private void testCase1_InsertRange1(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(4, sheet.getPoiSheet().getNumMergedRegions());
		
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);
		//right
		Ranges.range(sheet,"G1:G5").insert(Range.InsertShift.RIGHT,Range.InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","L2:N4","B10:D12","K10:M12"}, sheet);
		
		Ranges.range(sheet,"G1:H13").insert(Range.InsertShift.RIGHT,Range.InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","N2:P4","B10:D12","M10:O12"}, sheet);
		
		Ranges.range(sheet,"I5:J9").insert(Range.InsertShift.RIGHT,Range.InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","N2:P4","B10:D12","M10:O12"}, sheet);
		
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","N2:P4","B10:D12","M10:O12"}, sheet);
		
		//down
		Ranges.range(sheet,"A6:E6").insert(Range.InsertShift.DOWN,Range.InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","N2:P4","B11:D13","M10:O12"}, sheet);
		
		Ranges.range(sheet,"F6:H7").insert(Range.InsertShift.DOWN,Range.InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","N2:P4","B11:D13","M10:O12"}, sheet);
		
		
		//export and test again
		temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","N2:P4","B11:D13","M10:O12"}, sheet);
	}
	@Test
	public void testCase1_DeleteRange() throws IOException {
		testCase1_DeleteRange1(Util.loadBook(this,"book/434-case1.xlsx"));
		testCase1_DeleteRange1(mergeForCase1(Util.loadBook(this,"book/blank.xlsx")));
		testCase1_DeleteRange1(Util.loadBook(this,"book/434-case1.xls"));
		testCase1_DeleteRange1(mergeForCase1(Util.loadBook(this,"book/blank.xls")));
	}
	private void testCase1_DeleteRange1(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(4, sheet.getPoiSheet().getNumMergedRegions());
		
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);
		
		Ranges.range(sheet,"F1:G1").toColumnRange().delete(Range.DeleteShift.LEFT);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","I2:K4","B10:D12","I10:K12"}, sheet);
		
		Ranges.range(sheet,"J1").toColumnRange().delete(Range.DeleteShift.LEFT);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","I2:J4","B10:D12","I10:J12"}, sheet);
		
		Ranges.range(sheet,"I1").toColumnRange().delete(Range.DeleteShift.LEFT);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","I2:I4","B10:D12","I10:I12"}, sheet);
//		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","I2:I4","B10:D12","I10:I12"}, sheet);
	}
	
	
	
	@Test
	public void testBlank_MoveRange() throws IOException {
		testBlank_MoveRange(Util.loadBook(this,"book/blank.xlsx"));
		testBlank_MoveRange(Util.loadBook(this,"book/blank.xls"));
	}
	private void testBlank_MoveRange(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
		
		Ranges.range(sheet,"B2:D4").merge(false);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4"}, sheet);
		
		
		Ranges.range(sheet,"B1:D5").shift(1, 1);
		AssertUtil.assertMeregedRegion(new String[]{"C3:E5"}, sheet);
		
		Ranges.range(sheet,"B1:F7").shift(2, 2);
		AssertUtil.assertMeregedRegion(new String[]{"E5:G7"}, sheet);
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"E5:G7"}, sheet);
		
		Ranges.range(sheet,"E5:G7").shift(-3, -3);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4"}, sheet);

		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4"}, sheet);
	}
	
	
	@Test
	public void testCase1_MoveRange() throws IOException {
		
		testCase1_MoveRange1(Util.loadBook(this,"book/434-case1.xlsx"));
		testCase1_MoveRange1(mergeForCase1(Util.loadBook(this,"book/blank.xlsx")));
		testCase1_MoveRange1(Util.loadBook(this,"book/434-case1.xls"));
		testCase1_MoveRange1(mergeForCase1(Util.loadBook(this,"book/blank.xls")));
		
		testCase1_MoveRange2(Util.loadBook(this,"book/434-case1.xlsx"));
		testCase1_MoveRange2(mergeForCase1(Util.loadBook(this,"book/blank.xlsx")));
		testCase1_MoveRange2(Util.loadBook(this,"book/434-case1.xls"));
		testCase1_MoveRange2(mergeForCase1(Util.loadBook(this,"book/blank.xls")));
	}
	private void testCase1_MoveRange1(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);
		
		
		Ranges.range(sheet,"B1:D5").shift(1, 1);
		AssertUtil.assertMeregedRegion(new String[]{"C3:E5","K2:M4","B10:D12","K10:M12"}, sheet);
		
		Ranges.range(sheet,"B1:F7").shift(2, 2);
		AssertUtil.assertMeregedRegion(new String[]{"E5:G7","K2:M4","B10:D12","K10:M12"}, sheet);
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"E5:G7","K2:M4","B10:D12","K10:M12"}, sheet);
		
		Ranges.range(sheet,"E5:G7").shift(-3, -3);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);

		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);
	}
	
	private void testCase1_MoveRange2(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);
		
		
		Ranges.range(sheet,"K9:M13").shift(1, 1);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","L11:N13"}, sheet);
		
		Ranges.range(sheet,"K10:O14").shift(2, 2);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","N13:P15"}, sheet);
		
		//export and test again
		File temp = Setup.getTempFile();
		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","N13:P15"}, sheet);
		
		Ranges.range(sheet,"N13:Q15").shift(-3, -3);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);

		Exporters.getExporter().export(book, temp);
		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		AssertUtil.assertMeregedRegion(new String[]{"B2:D4","K2:M4","B10:D12","K10:M12"}, sheet);
	}

	
}
