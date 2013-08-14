package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.api.Importers;
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
	
	private static Book _workbook;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void setUp() throws Exception {
		final String filename = "book/shiftTest.xlsx";
		final InputStream is = PasteTest.class.getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
	}
	
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	@Test
	public void testShiftUpG4G6() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
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
	
	@Test
	public void testShiftDownE4E5() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
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
	
	@Test
	public void testShiftDownG3G5() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
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
	
	@Test
	public void testShiftLeftE3E5() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range range_E3E5 = Ranges.range(sheet1, "E3:E5");
		range_E3E5.delete(DeleteShift.LEFT);
		
		assertEquals("G3", Ranges.range(sheet1, "F3").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet1, "F4").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet1, "F5").getCellData().getEditText());

		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E4").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E5").getCellData().getType().ordinal(), 1E-8);
	}
	
	@Test
	public void testShiftRightE3E5() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
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
	
	@Test
	public void testDeleteColumnE() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
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
	
	@Test
	public void testDeleteRow345() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range range = Ranges.range(sheet1, "A3:A5");
		range.toRowRange().delete(DeleteShift.DEFAULT);
		
		assertEquals("C8", Ranges.range(sheet1, "C5").getCellData().getEditText());
		assertEquals("E6", Ranges.range(sheet1, "E3").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet1, "E4").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet1, "E5").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet1, "G5").getCellData().getEditText());
	}
	
	@Test
	public void testE3G5ShiftRow3Col3() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
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
