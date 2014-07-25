package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.AssertUtil;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.CellData.CellType;

/**
 * test
 * 1. shift
 * 2. insert
 * 3. delete
 * @author kuro
 *
 */
public class ShiftTest {
	
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
	
	@Test
	public void testDeleteAndInsertRowMerge2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testDeleteAndInsertRowMerge0(book);
	}
	
	@Test
	public void testDeleteAndInsertRowMerge2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testDeleteAndInsertRowMerge0(book);
	}
	
	@Test
	public void testDeleteAndInsertRow2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testDeleteAndInsertRow0(book, Setup.getTempFile());
	}
	
	@Test
	public void testDeleteAndInsertRow2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testDeleteAndInsertRow0(book, Setup.getTempFile());
	}
	
	@Test
	public void testDeleteAndInsertColumnBeforeAndAfterExport2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testDeleteAndInsertColumnBeforeAndAfterExport0(book, Setup.getTempFile());
	}
	
	@Test
	public void testDeleteAndInsertColumnBeforeAndAfterExport2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testDeleteAndInsertColumnBeforeAndAfterExport0(book, Setup.getTempFile());
	}
	
	@Test
	public void testDeleteAndInsertColumnBefore2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testDeleteAndInsertColumnBefore0(book, Setup.getTempFile());
	}
	
	@Test
	public void testDeleteAndInsertColumnBefore2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testDeleteAndInsertColumnBefore0(book, Setup.getTempFile());
	}

	@Test
	public void testShiftUpG4G6_2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testShiftUpG4G6(book);
	}
	
	@Test
	public void testShiftUpG4G6_2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testShiftUpG4G6(book);
	}
	
	@Test
	public void testShiftDownE4E5_2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testShiftDownE4E5(book);
	}
	
	@Test
	public void testShiftDownE4E5_2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testShiftDownE4E5(book);
	}
	
	@Test
	public void testShiftDownG3G5_2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testShiftDownG3G5(book);
	}
	
	@Test
	public void testShiftDownG3G5_2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testShiftDownG3G5(book);
	}
	
	@Test
	public void testShiftLeftE3E5_2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testShiftLeftE3E5(book);
	}
	
	@Test
	public void testShiftLeftE3E5_2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testShiftLeftE3E5(book);
	}
	
	@Test
	public void testShiftRightE3E5_2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testShiftRightE3E5(book);
	}
	
	@Test
	public void testShiftRightE3E5_2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testShiftRightE3E5(book);	
	}
	
	@Test
	public void testDeleteColumnE_2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testDeleteColumnE(book);
	}
	
	@Test
	public void testDeleteColumnE_2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testDeleteColumnE(book);
	}
	
	@Test
	public void testDeleteRow345_2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testDeleteRow345(book);
	}
	
	@Test
	public void testDeleteRow345_2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testDeleteRow345(book);
	}
	
	@Test
	public void testE3G5ShiftRow3Col3_2003() throws IOException {
		Book book = Util.loadBook("shiftTest.xls");
		testE3G5ShiftRow3Col3(book);
	}

	@Test
	public void testE3G5ShiftRow3Col3_2007() throws IOException {
		Book book = Util.loadBook("shiftTest.xlsx");
		testE3G5ShiftRow3Col3(book);
	}
	
	private void testDeleteAndInsertRowMerge0(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		
		Range rE3G5 = Ranges.range(sheet, "E3:G5");
		rE3G5.merge(false);
		
		Range rA4 = Ranges.range(sheet, "A4"); // E4, whole row cross merge area
		rA4.toRowRange().delete(DeleteShift.DEFAULT);
		
		AssertUtil.assertNotMergedRange(rE3G5); // should not be merge area anymore
		Range rE3G4 = Ranges.range(sheet, "E3:G4" );
		AssertUtil.assertMergedRange(rE3G4); // should be merge area
	}
	
	/**
	 * 1. delete row 3
	 * 1.1 validate
	 * 2. export
	 * 3. insert row 6
	 * 3.3 validate
	 * 4. Merge C4:E6
	 * 5. delete row 5 <- wrong in view, correct in model
	 * 5.5 validate
	 * 6. export
	 * 7. import & validate
	 */
	private void testDeleteAndInsertRow0(Book workbook, File outFile) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		// delete row 3
		Range row3 = Ranges.range(sheet, "A3");
		row3.toRowRange().delete(DeleteShift.DEFAULT);
		
		// validate
		assertEquals("E1", Ranges.range(sheet, "E1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "E2").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "E3").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "E4").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "E5").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet, "E6").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "E7").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C7").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "G7").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet, "G3").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "G4").getCellData().getEditText());
		
		Util.export(workbook, outFile);
		
		sheet = null;			// clean sheet
		row3 = null;			// clean range
		workbook = Util.loadBook(outFile);  // Import
		sheet = workbook.getSheet("Sheet1"); // get sheet
		
		// validate
		assertEquals("E1", Ranges.range(sheet, "E1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "E2").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "E3").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "E4").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "E5").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet, "E6").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "E7").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C7").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "G7").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet, "G3").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "G4").getCellData().getEditText());
		
		// insert row 6
		Range row6 = Ranges.range(sheet, "A6");
		row6.toRowRange().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		
		// validate
		assertEquals("E1", Ranges.range(sheet, "E1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "E2").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "E3").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "E4").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "E5").getCellData().getEditText());
		assertEquals(CellType.BLANK, Ranges.range(sheet, "E6").getCellData().getType());
		assertEquals("E7", Ranges.range(sheet, "E7").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "E8").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C8").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "G8").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet, "G3").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "G4").getCellData().getEditText());
		
		// Merge C4:E6
		Range rC4E6 = Ranges.range(sheet, "C4:E6");
		rC4E6.merge(false);
		AssertUtil.assertMergedRange(rC4E6);
		
		// delete row 5
		Range row5 = Ranges.range(sheet, "A5");
		row5.toRowRange().delete(DeleteShift.DEFAULT);
		
		// validate
		AssertUtil.assertNotMergedRange(rC4E6);
		
		assertEquals("E1", Ranges.range(sheet, "E1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "E2").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "E3").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "C4").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet, "E6").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "E7").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C7").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "G7").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet, "G3").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "G4").getCellData().getEditText());		
		
		Range rC4E5 = Ranges.range(sheet, "C4:E5");
		AssertUtil.assertMergedRange(rC4E5);
		
		Util.export(workbook, outFile);
		
		// import
		workbook = null; 		// clean work book
		sheet = null;			// clean sheet
		row3 = null;			// clean range
		workbook = Util.loadBook(outFile);  // Import
		sheet = workbook.getSheet("Sheet1"); // get sheet
		
		// validate
		AssertUtil.assertNotMergedRange(rC4E6);
		
		assertEquals("E1", Ranges.range(sheet, "E1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "E2").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "E3").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "C4").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet, "E6").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "E7").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C7").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "G7").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet, "G3").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "G4").getCellData().getEditText());		
		
		rC4E5 = Ranges.range(sheet, "C4:E5");
		AssertUtil.assertMergedRange(rC4E5);
	}
	
	/**
	 * 1. delete column D
	 * 2. export
	 * 3. edit cell E3 text A
	 * 4. edit cell D7 text B
	 * 5. insert row 5
	 * 6. export
	 * 7. import & validate
	 * 8. edit cell B6 text C
	 * 9. delete range B7:F8 (shift up)
	 * 10. export
	 * 11. delete range D3:E4 (shift left)
	 * 12. edit cell E4 text D
	 * 13. export
	 * 14. import & validate
	 */
	private void testDeleteAndInsertColumnBeforeAndAfterExport0(Book workbook, File outFile) throws IOException {
		
		Sheet sheet = workbook.getSheet("Sheet1");
		
		// 1. Delete column D & validate
		Range columnD = Ranges.range(sheet, "D1");
		columnD.toColumnRange().delete(DeleteShift.DEFAULT);
		
		assertEquals("E1", Ranges.range(sheet, "D1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "D2").getCellData().getEditText());
		assertEquals("E3", Ranges.range(sheet, "D3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "D4").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "D5").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "D6").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet, "D7").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "D8").getCellData().getEditText());
		
		// 2. export & validate
		Util.export(workbook, outFile);
		
		assertEquals("E1", Ranges.range(sheet, "D1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "D2").getCellData().getEditText());
		assertEquals("E3", Ranges.range(sheet, "D3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "D4").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "D5").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "D6").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet, "D7").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "D8").getCellData().getEditText());
		
		// 3. edit cell E3 text to A
		Ranges.range(sheet, "E3").setCellEditText("A");
		assertEquals("A", Ranges.range(sheet, "E3").getCellData().getEditText());
		// 4. edit cell D7 text B
		Ranges.range(sheet, "D7").setCellEditText("B");
		assertEquals("B", Ranges.range(sheet, "D7").getCellEditText());
		
		// 5. insert row 5
		Range row5 = Ranges.range(sheet, "A5");
		row5.toRowRange().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		
		assertEquals("E1", Ranges.range(sheet, "D1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "D2").getCellData().getEditText());
		assertEquals("E3", Ranges.range(sheet, "D3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "D4").getCellData().getEditText());
		assertEquals(CellType.BLANK, Ranges.range(sheet, "D5").getCellData().getType());
		assertEquals("E5", Ranges.range(sheet, "D6").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "D7").getCellData().getEditText());
		assertEquals("B", Ranges.range(sheet, "D8").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "D9").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C9").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "F9").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "F6").getCellData().getEditText());
		
		// 6. export & validate
		Util.export(workbook, outFile);
		
		assertEquals("A", Ranges.range(sheet, "E3").getCellEditText());
		assertEquals("B", Ranges.range(sheet, "D8").getCellEditText());
		
		assertEquals("E1", Ranges.range(sheet, "D1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "D2").getCellData().getEditText());
		assertEquals("E3", Ranges.range(sheet, "D3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "D4").getCellData().getEditText());
		assertEquals(CellType.BLANK, Ranges.range(sheet, "D5").getCellData().getType());
		assertEquals("E5", Ranges.range(sheet, "D6").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "D7").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "D9").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C9").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "F9").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "F6").getCellData().getEditText());
		
		// 7. import & validate
		workbook = null; 		// clean work book
		sheet = null;			// clean sheet
		workbook = Util.loadBook(outFile);  // Import
		sheet = workbook.getSheet("Sheet1"); // get sheet
		
		assertEquals("A", Ranges.range(sheet, "E3").getCellEditText());
		assertEquals("B", Ranges.range(sheet, "D8").getCellEditText());
		
		assertEquals("E1", Ranges.range(sheet, "D1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "D2").getCellData().getEditText());
		assertEquals("E3", Ranges.range(sheet, "D3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "D4").getCellData().getEditText());
		assertEquals(CellType.BLANK, Ranges.range(sheet, "D5").getCellData().getType());
		assertEquals("E5", Ranges.range(sheet, "D6").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "D7").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "D9").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C9").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "F9").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "F6").getCellData().getEditText());
		
		// 8. edit cell B6 text C
		Ranges.range(sheet,"B6").setCellEditText("C");
		assertEquals("C", Ranges.range(sheet, "B6").getCellEditText());
		
		// 9. delete range B7:F8 (shift up)
		Ranges.range(sheet,"B7:F8").delete(DeleteShift.UP);
		assertEquals("C8", Ranges.range(sheet, "C7").getCellEditText());
		assertEquals("E8", Ranges.range(sheet, "D7").getCellEditText());
		assertEquals("G8", Ranges.range(sheet, "F7").getCellEditText());
		
		// 10. export
		Util.export(workbook, outFile);
		
		assertEquals("A", Ranges.range(sheet, "E3").getCellEditText());
		assertEquals("C", Ranges.range(sheet, "B6").getCellEditText());
		
		assertEquals("E1", Ranges.range(sheet, "D1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "D2").getCellData().getEditText());
		assertEquals("E3", Ranges.range(sheet, "D3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "D4").getCellData().getEditText());
		assertEquals(CellType.BLANK, Ranges.range(sheet, "D5").getCellData().getType());
		assertEquals("E5", Ranges.range(sheet, "D6").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "D7").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C7").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "F7").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "F6").getCellData().getEditText());
		
		// 11. delete D3:E4 (shift left)
		Ranges.range(sheet,"D3:E4").delete(DeleteShift.LEFT);
		assertEquals("G3", Ranges.range(sheet, "D3").getCellEditText());
		assertEquals("G4", Ranges.range(sheet, "D4").getCellEditText());
		
		// 12. edit cell E4 text D
		Ranges.range(sheet,"E4").setCellEditText("D");
		assertEquals("D", Ranges.range(sheet, "E4").getCellEditText());
		
		assertEquals("E1", Ranges.range(sheet, "D1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "D2").getCellData().getEditText());
		assertEquals("G3", Ranges.range(sheet, "D3").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet, "D4").getCellData().getEditText());
		assertEquals(CellType.BLANK, Ranges.range(sheet, "D5").getCellData().getType());
		assertEquals("E5", Ranges.range(sheet, "D6").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "D7").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C7").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "F7").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "F6").getCellData().getEditText());
		assertEquals("C", Ranges.range(sheet, "B6").getCellData().getEditText());
		
		// 13. export
		Util.export(workbook, outFile);
		
		// 14. import & validate
		workbook = null; 		// clean work book
		sheet = null;			// clean sheet
		workbook = Util.loadBook(outFile);  // Import
		sheet = workbook.getSheet("Sheet1"); // get sheet
		
		assertEquals("D", Ranges.range(sheet, "E4").getCellEditText());
		
		assertEquals("E1", Ranges.range(sheet, "D1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "D2").getCellData().getEditText());
		assertEquals("G3", Ranges.range(sheet, "D3").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet, "D4").getCellData().getEditText());
		assertEquals(CellType.BLANK, Ranges.range(sheet, "D5").getCellData().getType());
		assertEquals("E5", Ranges.range(sheet, "D6").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "D7").getCellData().getEditText());
		assertEquals("C8", Ranges.range(sheet, "C7").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet, "F7").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet, "F6").getCellData().getEditText());
		assertEquals("C", Ranges.range(sheet, "B6").getCellData().getEditText());
	}
	
	/**
	 * test case step
	 * 1. Import shiftText.xlsx
	 * 2. Delete Column D
	 * 3. Validate
	 * 4. Export to book/test.xlsx
	 * 5. Import book/test.xlsx
	 * 6. Validate
	 * 7. InsertColumn D
	 * 8. Validate
	 * 9. Export to book/test.xlsx
	 * 10. Import
	 * 11. Validate
	 */
	private void testDeleteAndInsertColumnBefore0(Book workbook, File outFile) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		
		// 2. Delete column D
		Range columnD = Ranges.range(sheet, "D1");
		columnD.toColumnRange().delete(DeleteShift.DEFAULT);
		
		// 3. Validate
		assertEquals("E1", Ranges.range(sheet, "D1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "D2").getCellData().getEditText());
		assertEquals("E3", Ranges.range(sheet, "D3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "D4").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "D5").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "D6").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet, "D7").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "D8").getCellData().getEditText());
		
		// 4. Export to book/test.xlsx
		Util.export(workbook, outFile);
		
		// 5. import
		workbook = null; 		// clean work book
		sheet = null;			// clean sheet
		columnD = null;			// clean range
		workbook = Util.loadBook(outFile);  // Import
		sheet = workbook.getSheet("Sheet1"); // get sheet
		
		// 6. Validate
		assertEquals("E1", Ranges.range(sheet, "D1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "D2").getCellData().getEditText());
		assertEquals("E3", Ranges.range(sheet, "D3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "D4").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "D5").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "D6").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet, "D7").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "D8").getCellData().getEditText());
		
		// 7. Insert column D
		columnD = Ranges.range(sheet, "D1");
		columnD.toColumnRange().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		
		// 8. Validate
		assertEquals("E1", Ranges.range(sheet, "E1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "E2").getCellData().getEditText());
		assertEquals("E3", Ranges.range(sheet, "E3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "E4").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "E5").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "E6").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet, "E7").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "E8").getCellData().getEditText());
		
		// 9. export
		Util.export(workbook, outFile);
		
		// 10. import

		sheet = null;			// clean sheet
		columnD = null;			// clean range
		workbook = Util.loadBook(outFile);  // Import
		sheet = workbook.getSheet("Sheet1"); // get sheet
		
		// 11. validate
		assertEquals("E1", Ranges.range(sheet, "E1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet, "E2").getCellData().getEditText());
		assertEquals("E3", Ranges.range(sheet, "E3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet, "E4").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet, "E5").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet, "E6").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet, "E7").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet, "E8").getCellData().getEditText());
	}

	private void testShiftUpG4G6(Book workbook) {
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range_G4G6 = Ranges.range(sheet1, "G4:G6");
		range_G4G6.delete(DeleteShift.UP);
		
		assertEquals("G3", Ranges.range(sheet1, "G3").getCellData().getEditText());
		
		// validate shift up
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G4").getCellData().getType().ordinal(), 1E-8);
		assertEquals("G8", Ranges.range(sheet1, "G5").getCellData().getEditText());
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G6").getCellData().getType().ordinal(), 1E-8);
		
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G8").getCellData().getType().ordinal(), 1E-8);
	}
	
	private void testShiftDownE4E5(Book workbook) {
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range_E4E5 = Ranges.range(sheet1, "E4:E5");
		range_E4E5.insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_NONE);
		
		// validate shift down
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E4").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E5").getCellData().getType().ordinal(), 1E-8);
		
		assertEquals("E4", Ranges.range(sheet1, "E6").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet1, "E7").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet1, "E8").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet1, "E9").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet1, "E10").getCellData().getEditText());
	}
	
	private void testShiftDownG3G5(Book workbook) {
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range_G3G5 = Ranges.range(sheet1, "G3:G5");
		range_G3G5.insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_NONE);
		
		// validate shift up
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G4").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G5").getCellData().getType().ordinal(), 1E-8);
		
		assertEquals("G3", Ranges.range(sheet1, "G6").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet1, "G7").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet1, "G8").getCellData().getEditText());
		
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G9").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G10").getCellData().getType().ordinal(), 1E-8);
	}
	
	private void testShiftLeftE3E5(Book workbook) {
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range_E3E5 = Ranges.range(sheet1, "E3:E5");
		range_E3E5.delete(DeleteShift.LEFT);
		
		assertEquals("G3", Ranges.range(sheet1, "F3").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet1, "F4").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet1, "F5").getCellData().getEditText());

		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E4").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E5").getCellData().getType().ordinal(), 1E-8);
	}
	
	private void testShiftRightE3E5(Book workbook) {
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range_E3E5 = Ranges.range(sheet1, "E3:E5");
		range_E3E5.insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_NONE);
		
		assertEquals("E3", Ranges.range(sheet1, "F3").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet1, "F4").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet1, "F5").getCellData().getEditText());

		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E4").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E5").getCellData().getType().ordinal(), 1E-8);
		
		assertEquals("G3", Ranges.range(sheet1, "H3").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet1, "H4").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet1, "H5").getCellData().getEditText());
	}
	
	private void testDeleteColumnE(Book workbook) {
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range_E = Ranges.range(sheet1, "E3");
		range_E.toColumnRange().delete(DeleteShift.DEFAULT);
		
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E1").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E2").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E4").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E5").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E8").getCellData().getType().ordinal(), 1E-8);
		
		assertEquals("G3", Ranges.range(sheet1, "F3").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet1, "F4").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet1, "F5").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet1, "F8").getCellData().getEditText());
	}
	
	private void testDeleteRow345(Book workbook) {
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range = Ranges.range(sheet1, "A3:A5");
		range.toRowRange().delete(DeleteShift.DEFAULT);
		
		assertEquals("C8", Ranges.range(sheet1, "C5").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet1, "E3").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet1, "E4").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet1, "E5").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet1, "G5").getCellData().getEditText());
	}
	
	private void testE3G5ShiftRow3Col3(Book workbook) {
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range = Ranges.range(sheet1, "E3:G5");
		range.shift(3, 3);
		assertEquals("E3", Ranges.range(sheet1, "H6").getCellData().getEditText());
		assertEquals("E4", Ranges.range(sheet1, "H7").getCellData().getEditText());
		assertEquals("E5", Ranges.range(sheet1, "H8").getCellData().getEditText());
		assertEquals("G3", Ranges.range(sheet1, "J6").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet1, "J7").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet1, "J8").getCellData().getEditText());
	}

}
