package org.zkoss.zss.api.impl.formula;

import org.junit.*;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.util.*;
import org.zkoss.zss.Setup;

import static org.junit.Assert.*;

public class AreaReferenceTest {

    @BeforeClass
    static public void beforeClass() {
        Setup.touch();
    }

    @Test
    public void testAreaRef1() {
        AreaReference ref = new AreaReference("$A$1:$B$2");
        assertFalse("Two cells expected", ref.isSingleCell());
        CellReference cf = ref.getFirstCell();
        assertEquals("row is 4", 0, cf.getRow());
        assertEquals("col is 1", 0, cf.getCol());
        assertTrue("row is abs",cf.isRowAbsolute());
        assertTrue("col is abs",cf.isColAbsolute());
        assertEquals("string is $A$1", "$A$1", cf.formatAsString());

        cf = ref.getLastCell();
        assertEquals("row is 4", 1, cf.getRow());
        assertEquals("col is 1", 1, cf.getCol());
        assertTrue("row is abs",cf.isRowAbsolute());
        assertTrue("col is abs",cf.isColAbsolute());
        assertEquals("string is $B$2", "$B$2", cf.formatAsString());

        CellReference[] refs = ref.getAllReferencedCells();
        assertEquals(4, refs.length);

        assertEquals(0, refs[0].getRow());
        assertEquals(0, refs[0].getCol());
        assertNull(refs[0].getSheetName());

        assertEquals(0, refs[1].getRow());
        assertEquals(1, refs[1].getCol());
        assertNull(refs[1].getSheetName());

        assertEquals(1, refs[2].getRow());
        assertEquals(0, refs[2].getCol());
        assertNull(refs[2].getSheetName());

        assertEquals(1, refs[3].getRow());
        assertEquals(1, refs[3].getCol());
        assertNull(refs[3].getSheetName());
    }

    @Test
    public void testReferenceWithSheet() {
        AreaReference ar;

        ar = new AreaReference("Tabelle1!B5:B5");
        assertTrue(ar.isSingleCell());
//        TestCellReference.confirmCell(ar.getFirstCell(), "Tabelle1", 4, 1, false, false, "Tabelle1!B5");

        assertEquals(1, ar.getAllReferencedCells().length);


        ar = new AreaReference("Tabelle1!$B$5:$B$7");
        assertFalse(ar.isSingleCell());

//        TestCellReference.confirmCell(ar.getFirstCell(), "Tabelle1", 4, 1, true, true, "Tabelle1!$B$5");
//        TestCellReference.confirmCell(ar.getLastCell(), "Tabelle1", 6, 1, true, true, "Tabelle1!$B$7");

        // And all that make it up
        CellReference[] allCells = ar.getAllReferencedCells();
        assertEquals(3, allCells.length);
//        TestCellReference.confirmCell(allCells[0], "Tabelle1", 4, 1, true, true, "Tabelle1!$B$5");
//        TestCellReference.confirmCell(allCells[1], "Tabelle1", 5, 1, true, true, "Tabelle1!$B$6");
//        TestCellReference.confirmCell(allCells[2], "Tabelle1", 6, 1, true, true, "Tabelle1!$B$7");
    }

    @Test
    public void testContiguousReferences() {
        String refSimple = "$C$10:$C$10";
        String ref2D = "$C$10:$D$11";
        String refDCSimple = "$C$10:$C$10,$D$12:$D$12,$E$14:$E$14";
        String refDC2D = "$C$10:$C$11,$D$12:$D$12,$E$14:$E$20";
        String refDC3D = "Tabelle1!$C$10:$C$14,Tabelle1!$D$10:$D$12";

        // Check that we detect as contiguous properly
        assertTrue(AreaReference.isContiguous(refSimple));
        assertTrue(AreaReference.isContiguous(ref2D));
        assertFalse(AreaReference.isContiguous(refDCSimple));
        assertFalse(AreaReference.isContiguous(refDC2D));
        assertFalse(AreaReference.isContiguous(refDC3D));

        // Check we can only create contiguous entries
        new AreaReference(refSimple);
        new AreaReference(ref2D);
        try {
            new AreaReference(refDCSimple);
            fail("expected IllegalArgumentException");
        } catch(IllegalArgumentException e) {
            // expected during successful test
        }
        try {
            new AreaReference(refDC2D);
            fail("expected IllegalArgumentException");
        } catch(IllegalArgumentException e) {
            // expected during successful test
        }
        try {
            new AreaReference(refDC3D);
            fail("expected IllegalArgumentException");
        } catch(IllegalArgumentException e) {
            // expected during successful test
        }

        // Test that we split as expected
        AreaReference[] refs;

        refs = AreaReference.generateContiguous( refSimple);
        assertEquals(1, refs.length);
        assertTrue(refs[0].isSingleCell());
        assertEquals("$C$10", refs[0].formatAsString());

        refs = AreaReference.generateContiguous( ref2D);
        assertEquals(1, refs.length);
        assertFalse(refs[0].isSingleCell());
        assertEquals("$C$10:$D$11", refs[0].formatAsString());

        refs = AreaReference.generateContiguous( refDCSimple);
        assertEquals(3, refs.length);
        assertTrue(refs[0].isSingleCell());
        assertTrue(refs[1].isSingleCell());
        assertTrue(refs[2].isSingleCell());
        assertEquals("$C$10", refs[0].formatAsString());
        assertEquals("$D$12", refs[1].formatAsString());
        assertEquals("$E$14", refs[2].formatAsString());

        refs = AreaReference.generateContiguous( refDC2D);
        assertEquals(3, refs.length);
        assertFalse(refs[0].isSingleCell());
        assertTrue(refs[1].isSingleCell());
        assertFalse(refs[2].isSingleCell());
        assertEquals("$C$10:$C$11", refs[0].formatAsString());
        assertEquals("$D$12", refs[1].formatAsString());
        assertEquals("$E$14:$E$20", refs[2].formatAsString());

        refs = AreaReference.generateContiguous( refDC3D);
        assertEquals(2, refs.length);
        assertFalse(refs[0].isSingleCell());
        assertFalse(refs[0].isSingleCell());
        assertEquals("Tabelle1!$C$10:$C$14", refs[0].formatAsString());
        assertEquals("Tabelle1!$D$10:$D$12", refs[1].formatAsString());
        assertEquals("Tabelle1", refs[0].getFirstCell().getSheetName());
        assertEquals("Tabelle1", refs[0].getLastCell().getSheetName());
        assertEquals("Tabelle1", refs[1].getFirstCell().getSheetName());
        assertEquals("Tabelle1", refs[1].getLastCell().getSheetName());
    }

    @Test
    public void testSpecialSheetNames() {
        AreaReference ar;
        ar = new AreaReference("'Sheet A'!A1:A1");
        confirmAreaSheetName(ar, "Sheet A", "'Sheet A'!A1");

        ar = new AreaReference("'Hey! Look Here!'!A1:A1");
        confirmAreaSheetName(ar, "Hey! Look Here!", "'Hey! Look Here!'!A1");

        ar = new AreaReference("'O''Toole'!A1:B2");
        confirmAreaSheetName(ar, "O'Toole", "'O''Toole'!A1:B2");

        //sheet name can't start or end with a single quote
        ar = new AreaReference("'one:many'!A1:B2");
        confirmAreaSheetName(ar, "one:many", "one:many!A1:B2");
    }

    @Test
    public void testWholeColumnRefs() {
        confirmWholeColumnRef("A:A", 0, 0, false, false);
        confirmWholeColumnRef("$C:D", 2, 3, true, false);
        confirmWholeColumnRef("AD:$AE", 29, 30, false, true);
    }


    private static void confirmAreaSheetName(AreaReference ar, String sheetName, String expectedFullText) {
        CellReference[] cells = ar.getAllReferencedCells();
        assertEquals(sheetName, cells[0].getSheetName());
        assertEquals(expectedFullText, ar.formatAsString());
    }

    private static void confirmWholeColumnRef(String ref, int firstCol, int lastCol, boolean firstIsAbs, boolean lastIsAbs) {
        AreaReference ar = new AreaReference(ref);
        confirmCell(ar.getFirstCell(), 0, firstCol, true, firstIsAbs);
        confirmCell(ar.getLastCell(), SpreadsheetVersion.EXCEL2007.getLastRowIndex(), lastCol, true, lastIsAbs);
    }


    private static void confirmCell(CellReference cell, int row, int col, boolean isRowAbs,
                                    boolean isColAbs) {
        assertEquals(row, cell.getRow());
        assertEquals(col, cell.getCol());
        assertEquals(isRowAbs, cell.isRowAbsolute());
        assertEquals(isColAbs, cell.isColAbsolute());
    }

    @Test
    public void testWholeRow(){
        AreaReference ref = new AreaReference("1:3");
        assertEquals("1:3", ref.formatAsString());
    }

    @Test
    public void testWholeColumn(){
        AreaReference ref = new AreaReference("A:B");
        Assert.assertTrue(ref.isWholeColumnReference());
        assertEquals("A:B", ref.formatAsString());
    }
}
