package org.zkoss.zss.model;

import java.util.List;
import java.util.Locale;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.PasteOption;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.PasteOption.PasteType;
import org.zkoss.zss.model.impl.BookImpl;
import org.zkoss.zss.model.impl.SheetImpl;

public class ModelCopyTest {
	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}
	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	protected SSheet initialDataGrid(SSheet sheet){
		return sheet;
	}
	
	@Test 
	public void testCopySimple(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue("A");
		sheet1.getCell("A2").setValue(13);
		sheet1.getCell("A3").setValue("=A2");
		sheet1.getCell("A4").setValue("=SUM(A2:A3)");
		
		CellRegion finalRegion = sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("B2"), null);
		Assert.assertEquals("B2:B5", finalRegion.getReferenceString());
		
		Assert.assertEquals("A", sheet1.getCell("B2").getValue());
		Assert.assertEquals(13D, sheet1.getCell("B3").getValue());
		Assert.assertEquals("B3", sheet1.getCell("B4").getFormulaValue());
		Assert.assertEquals("SUM(B3:B4)", sheet1.getCell("B5").getFormulaValue());
		
		Assert.assertEquals(13D, sheet1.getCell("B4").getValue());
		Assert.assertEquals(26D, sheet1.getCell("B5").getValue());
		
		sheet1.getCell("B3").setValue(10);
		Assert.assertEquals(10D, sheet1.getCell("B4").getValue());
		Assert.assertEquals(20D, sheet1.getCell("B5").getValue());
		
		//multiple row
		finalRegion = sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("C2:C9"), null);
		Assert.assertEquals("C2:C9", finalRegion.getReferenceString());
		
		Assert.assertEquals("A", sheet1.getCell("C2").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C3").getValue());
		Assert.assertEquals("C3", sheet1.getCell("C4").getFormulaValue());
		Assert.assertEquals("SUM(C3:C4)", sheet1.getCell("C5").getFormulaValue());
		Assert.assertEquals("A", sheet1.getCell("C6").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C7").getValue());
		Assert.assertEquals("C7", sheet1.getCell("C8").getFormulaValue());
		Assert.assertEquals("SUM(C7:C8)", sheet1.getCell("C9").getFormulaValue());
		
		sheet1.getCell("C3").setValue(12);
		Assert.assertEquals(12D, sheet1.getCell("C4").getValue());
		Assert.assertEquals(24D, sheet1.getCell("C5").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C8").getValue());
		Assert.assertEquals(26D, sheet1.getCell("C9").getValue());
		
		//multiple row/column
		finalRegion = sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("C2:D9"), null);
		Assert.assertEquals("C2:D9", finalRegion.getReferenceString());
		Assert.assertEquals("A", sheet1.getCell("C2").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C3").getValue());
		Assert.assertEquals("C3", sheet1.getCell("C4").getFormulaValue());
		Assert.assertEquals("SUM(C3:C4)", sheet1.getCell("C5").getFormulaValue());
		Assert.assertEquals("A", sheet1.getCell("C6").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C7").getValue());
		Assert.assertEquals("C7", sheet1.getCell("C8").getFormulaValue());
		Assert.assertEquals("SUM(C7:C8)", sheet1.getCell("C9").getFormulaValue());
		//
		Assert.assertEquals("A", sheet1.getCell("D2").getValue());
		Assert.assertEquals(13D, sheet1.getCell("D3").getValue());
		Assert.assertEquals("D3", sheet1.getCell("D4").getFormulaValue());
		Assert.assertEquals("SUM(D3:D4)", sheet1.getCell("D5").getFormulaValue());
		Assert.assertEquals("A", sheet1.getCell("D6").getValue());
		Assert.assertEquals(13D, sheet1.getCell("D7").getValue());
		Assert.assertEquals("D7", sheet1.getCell("D8").getFormulaValue());
		Assert.assertEquals("SUM(D7:D8)", sheet1.getCell("D9").getFormulaValue());
		
		
		sheet1.getCell("C3").setValue(12);
		Assert.assertEquals(12D, sheet1.getCell("C4").getValue());
		Assert.assertEquals(24D, sheet1.getCell("C5").getValue());
		Assert.assertEquals(13D, sheet1.getCell("C8").getValue());
		Assert.assertEquals(26D, sheet1.getCell("C9").getValue());		
		
		sheet1.getCell("D7").setValue(12);
		Assert.assertEquals(13D, sheet1.getCell("D4").getValue());
		Assert.assertEquals(26D, sheet1.getCell("D5").getValue());
		Assert.assertEquals(12D, sheet1.getCell("D8").getValue());
		Assert.assertEquals(24D, sheet1.getCell("D9").getValue());
	}
	
	@Test 
	public void testCopyMerge(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("B2").setValue("A");
		sheet1.getCell("B3").setValue("B");
		
		sheet1.addMergedRegion(new CellRegion("B2:C2"));
		sheet1.addMergedRegion(new CellRegion("B3:C4"));
		
		sheet1.pasteCell(new SheetRegion(sheet1, "B2:C4"), new CellRegion("D4"), null);
		
		Assert.assertEquals("A", sheet1.getCell("D4").getValue());
		Assert.assertEquals("B", sheet1.getCell("D5").getValue());
		
		List<CellRegion> overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("D4:E4"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("D4:E4", overlaps.get(0).getReferenceString());
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("D5:E6"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("D5:E6", overlaps.get(0).getReferenceString());
		
		//repeat x
		sheet1.addMergedRegion(new CellRegion("K4:K6"));
		sheet1.pasteCell(new SheetRegion(sheet1, "B2:C4"), new CellRegion("F4:K6"), null);
		Assert.assertEquals("A", sheet1.getCell("F4").getValue());
		Assert.assertEquals("B", sheet1.getCell("F5").getValue());
		Assert.assertEquals("A", sheet1.getCell("H4").getValue());
		Assert.assertEquals("B", sheet1.getCell("H5").getValue());
		Assert.assertEquals("A", sheet1.getCell("J4").getValue());
		Assert.assertEquals("B", sheet1.getCell("J5").getValue());
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("F4:G4"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("F4:G4", overlaps.get(0).getReferenceString());
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("F5:G6"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("F5:G6", overlaps.get(0).getReferenceString());
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("H4:I4"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("H4:I4", overlaps.get(0).getReferenceString());
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("H5:I6"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("H5:I6", overlaps.get(0).getReferenceString());
		
		//K4:K6 was unmerged
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("J4:K4"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("J4:K4", overlaps.get(0).getReferenceString());
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("J5:K6"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("J5:K6", overlaps.get(0).getReferenceString());
		
		//test cut
		PasteOption op = new PasteOption();
		op.setCut(true);
		sheet1.pasteCell(new SheetRegion(sheet1, "B2:C4"), new CellRegion("B5"), op);
		Assert.assertEquals("A", sheet1.getCell("B5").getValue());
		Assert.assertEquals("B", sheet1.getCell("B6").getValue());
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("B5:C5"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("B5:C5", overlaps.get(0).getReferenceString());
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("B6:C7"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("B6:C7", overlaps.get(0).getReferenceString());
		
		Assert.assertEquals(0, sheet1.getOverlapsMergedRegions(new CellRegion("B2:C2"),false).size());
		Assert.assertEquals(0, sheet1.getOverlapsMergedRegions(new CellRegion("B3:C4"),false).size());
		
		//test cut and overlap
		sheet1.pasteCell(new SheetRegion(sheet1, "B5:C7"), new CellRegion("A4"), op);
		Assert.assertEquals("A", sheet1.getCell("A4").getValue());
		Assert.assertEquals("B", sheet1.getCell("A5").getValue());
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("A4:B4"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("A4:B4", overlaps.get(0).getReferenceString());
		
		overlaps = sheet1.getOverlapsMergedRegions(new CellRegion("A5:B6"),false);
		Assert.assertEquals(1, overlaps.size());
		Assert.assertEquals("A5:B6", overlaps.get(0).getReferenceString());
		
	}
	
	@Test 
	public void testCopyColumnWidth(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getColumn(0).setWidth(10);
		sheet1.getColumn(1).setWidth(20);
		sheet1.getColumn(2).setWidth(30);
		PasteOption opt = new PasteOption();
		opt.setPasteType(PasteType.COLUMN_WIDTHS);
		sheet1.pasteCell(new SheetRegion(sheet1,new CellRegion("A1:C1")), new CellRegion("B1:G1"), opt);
		
		Assert.assertEquals(10, sheet1.getColumn(1).getWidth());
		Assert.assertEquals(20, sheet1.getColumn(2).getWidth());
		Assert.assertEquals(30, sheet1.getColumn(3).getWidth());
		Assert.assertEquals(10, sheet1.getColumn(4).getWidth());
		Assert.assertEquals(20, sheet1.getColumn(5).getWidth());
		Assert.assertEquals(30, sheet1.getColumn(6).getWidth());
	}
	
	@Test 
	public void testCopyTranspose(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("D4").setValue("Sales by Region");
		sheet1.getCell("D5").setValue("Q1");
		sheet1.getCell("D6").setValue("Q2");
		sheet1.getCell("D7").setValue("Q3");
		sheet1.getCell("D8").setValue("Q4");
		sheet1.getCell("E4").setValue("Eurpoe");
		sheet1.getCell("E5").setValue(10);
		sheet1.getCell("E6").setValue(30);
		sheet1.getCell("E7").setValue(50);
		sheet1.getCell("E8").setValue(70);
		sheet1.getCell("F4").setValue("Asia");
		sheet1.getCell("F5").setValue(20);
		sheet1.getCell("F6").setValue(40);
		sheet1.getCell("F7").setValue(60);
		sheet1.getCell("F8").setValue(80);
		sheet1.getCell("G4").setValue("Total");
		sheet1.getCell("G5").setValue("=SUM(E5:F5)");
		sheet1.getCell("G6").setValue("=SUM(E6:F6)");
		sheet1.getCell("G7").setValue("=SUM(E7:F7)");
		sheet1.getCell("G8").setValue("=SUM(E8:F8)");
		
		Assert.assertEquals(30D, sheet1.getCell("G5").getValue());
		Assert.assertEquals(70D, sheet1.getCell("G6").getValue());
		Assert.assertEquals(110D, sheet1.getCell("G7").getValue());
		Assert.assertEquals(150D, sheet1.getCell("G8").getValue());
		
		
		//paste transpose
		PasteOption option = new PasteOption();
		option.setTranspose(true);
		sheet1.pasteCell(new SheetRegion(sheet1,new CellRegion("D4:G8")), new CellRegion("E11"), option);
		Assert.assertEquals("Sales by Region", sheet1.getCell("E11").getValue());
		Assert.assertEquals("Q1",sheet1.getCell("F11").getValue());
		Assert.assertEquals("Q2",sheet1.getCell("G11").getValue());
		Assert.assertEquals("Q3",sheet1.getCell("H11").getValue());
		Assert.assertEquals("Q4",sheet1.getCell("I11").getValue());
		
		Assert.assertEquals("Eurpoe", sheet1.getCell("E12").getValue());
		Assert.assertEquals(10D,sheet1.getCell("F12").getValue());
		Assert.assertEquals(30D,sheet1.getCell("G12").getValue());
		Assert.assertEquals(50D,sheet1.getCell("H12").getValue());
		Assert.assertEquals(70D,sheet1.getCell("I12").getValue());
		
		Assert.assertEquals("Asia", sheet1.getCell("E13").getValue());
		Assert.assertEquals(20D,sheet1.getCell("F13").getValue());
		Assert.assertEquals(40D,sheet1.getCell("G13").getValue());
		Assert.assertEquals(60D,sheet1.getCell("H13").getValue());
		Assert.assertEquals(80D,sheet1.getCell("I13").getValue());
		
		Assert.assertEquals("Total", sheet1.getCell("E14").getValue());
		
		Assert.assertEquals("SUM(F12:F13)",sheet1.getCell("F14").getFormulaValue());
		Assert.assertEquals("SUM(G12:G13)",sheet1.getCell("G14").getFormulaValue());
		Assert.assertEquals("SUM(H12:H13)",sheet1.getCell("H14").getFormulaValue());
		Assert.assertEquals("SUM(I12:I13)",sheet1.getCell("I14").getFormulaValue());
		
		Assert.assertEquals(30D, sheet1.getCell("F14").getValue());
		Assert.assertEquals(70D, sheet1.getCell("G14").getValue());
		Assert.assertEquals(110D, sheet1.getCell("H14").getValue());
		Assert.assertEquals(150D, sheet1.getCell("I14").getValue());
		
		
		//paste transpose
		sheet1.clearCell(new CellRegion("D11:H14"));
		
		sheet1.pasteCell(new SheetRegion(sheet1,new CellRegion("D4:G8")), new CellRegion("E11:I18"), option);
		Assert.assertEquals("Sales by Region", sheet1.getCell("E11").getValue());
		Assert.assertEquals("Q1",sheet1.getCell("F11").getValue());
		Assert.assertEquals("Q2",sheet1.getCell("G11").getValue());
		Assert.assertEquals("Q3",sheet1.getCell("H11").getValue());
		Assert.assertEquals("Q4",sheet1.getCell("I11").getValue());
		
		Assert.assertEquals("Eurpoe", sheet1.getCell("E12").getValue());
		Assert.assertEquals(10D,sheet1.getCell("F12").getValue());
		Assert.assertEquals(30D,sheet1.getCell("G12").getValue());
		Assert.assertEquals(50D,sheet1.getCell("H12").getValue());
		Assert.assertEquals(70D,sheet1.getCell("I12").getValue());
		
		Assert.assertEquals("Asia", sheet1.getCell("E13").getValue());
		Assert.assertEquals(20D,sheet1.getCell("F13").getValue());
		Assert.assertEquals(40D,sheet1.getCell("G13").getValue());
		Assert.assertEquals(60D,sheet1.getCell("H13").getValue());
		Assert.assertEquals(80D,sheet1.getCell("I13").getValue());
		
		Assert.assertEquals("Total", sheet1.getCell("E14").getValue());
		
		Assert.assertEquals("SUM(F12:F13)",sheet1.getCell("F14").getFormulaValue());
		Assert.assertEquals("SUM(G12:G13)",sheet1.getCell("G14").getFormulaValue());
		Assert.assertEquals("SUM(H12:H13)",sheet1.getCell("H14").getFormulaValue());
		Assert.assertEquals("SUM(I12:I13)",sheet1.getCell("I14").getFormulaValue());
		
		Assert.assertEquals(30D, sheet1.getCell("F14").getValue());
		Assert.assertEquals(70D, sheet1.getCell("G14").getValue());
		Assert.assertEquals(110D, sheet1.getCell("H14").getValue());
		Assert.assertEquals(150D, sheet1.getCell("I14").getValue());
		
		
		//the repeat 2
		Assert.assertEquals("Sales by Region", sheet1.getCell("E15").getValue());
		Assert.assertEquals("Q1",sheet1.getCell("F15").getValue());
		Assert.assertEquals("Q2",sheet1.getCell("G15").getValue());
		Assert.assertEquals("Q3",sheet1.getCell("H15").getValue());
		Assert.assertEquals("Q4",sheet1.getCell("I15").getValue());
		
		Assert.assertEquals("Eurpoe", sheet1.getCell("E16").getValue());
		Assert.assertEquals(10D,sheet1.getCell("F16").getValue());
		Assert.assertEquals(30D,sheet1.getCell("G16").getValue());
		Assert.assertEquals(50D,sheet1.getCell("H16").getValue());
		Assert.assertEquals(70D,sheet1.getCell("I16").getValue());
		
		Assert.assertEquals("Asia", sheet1.getCell("E17").getValue());
		Assert.assertEquals(20D,sheet1.getCell("F17").getValue());
		Assert.assertEquals(40D,sheet1.getCell("G17").getValue());
		Assert.assertEquals(60D,sheet1.getCell("H17").getValue());
		Assert.assertEquals(80D,sheet1.getCell("I17").getValue());
		
		Assert.assertEquals("Total", sheet1.getCell("E18").getValue());
		
		Assert.assertEquals("SUM(F16:F17)",sheet1.getCell("F18").getFormulaValue());
		Assert.assertEquals("SUM(G16:G17)",sheet1.getCell("G18").getFormulaValue());
		Assert.assertEquals("SUM(H16:H17)",sheet1.getCell("H18").getFormulaValue());
		Assert.assertEquals("SUM(I16:I17)",sheet1.getCell("I18").getFormulaValue());
		
		Assert.assertEquals(30D, sheet1.getCell("F18").getValue());
		Assert.assertEquals(70D, sheet1.getCell("G18").getValue());
		Assert.assertEquals(110D, sheet1.getCell("H18").getValue());
		Assert.assertEquals(150D, sheet1.getCell("I18").getValue());
	}
	
	public static void dump(SBook book){
		StringBuilder builder = new StringBuilder();
		((BookImpl)book).dump(builder);
		System.out.println(builder.toString());
	}
	
	@Test 
	public void testCopyTransposeMerge(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("J4").setValue("Merge");
		sheet1.getCell("K5").setValue("K5:L5");
		sheet1.getCell("K6").setValue("K6:K7");
		sheet1.getCell("L6").setValue("L6:N7");
		sheet1.addMergedRegion(new CellRegion("K5:L5"));
		sheet1.addMergedRegion(new CellRegion("K6:K7"));
		sheet1.addMergedRegion(new CellRegion("L6:N7"));
		
		
		
		PasteOption option = new PasteOption();
		option.setTranspose(true);
		sheet1.pasteCell(new SheetRegion(sheet1,new CellRegion("J4:N7")), new CellRegion("K10"), option);
		Assert.assertEquals("Merge", sheet1.getCell("K10").getValue());
		
		Assert.assertEquals("K5:L5", sheet1.getCell("L11").getValue());
		Assert.assertEquals("K6:K7", sheet1.getCell("M11").getValue());
		Assert.assertEquals("L6:N7", sheet1.getCell("M12").getValue());
		
		Assert.assertEquals("L11:L12", sheet1.getMergedRegion("L11").getReferenceString());
		Assert.assertEquals("M11:N11", sheet1.getMergedRegion("M11").getReferenceString());
		Assert.assertEquals("M12:N14", sheet1.getMergedRegion("M12").getReferenceString());
		Assert.assertEquals(3, sheet1.getOverlapsMergedRegions(new CellRegion("K10:N14"), false).size());
		
		Assert.assertEquals(6, sheet1.getNumOfMergedRegion());
		sheet1.removeMergedRegion(new CellRegion("K10:N14"), true);
		Assert.assertEquals(3,sheet1.getNumOfMergedRegion());
		
		//copy repeat
		sheet1.pasteCell(new SheetRegion(sheet1,new CellRegion("J4:N7")), new CellRegion("K10:N19"), option);
		Assert.assertEquals("Merge", sheet1.getCell("K10").getValue());
		Assert.assertEquals("K5:L5", sheet1.getCell("L11").getValue());
		Assert.assertEquals("K6:K7", sheet1.getCell("M11").getValue());
		Assert.assertEquals("L6:N7", sheet1.getCell("M12").getValue());
		Assert.assertEquals("L11:L12", sheet1.getMergedRegion("L11").getReferenceString());
		Assert.assertEquals("M11:N11", sheet1.getMergedRegion("M11").getReferenceString());
		Assert.assertEquals("M12:N14", sheet1.getMergedRegion("M12").getReferenceString());
		Assert.assertEquals(3, sheet1.getOverlapsMergedRegions(new CellRegion("K10:N14"), false).size());
		Assert.assertEquals("Merge", sheet1.getCell("K15").getValue());
		Assert.assertEquals("K5:L5", sheet1.getCell("L16").getValue());
		Assert.assertEquals("K6:K7", sheet1.getCell("M16").getValue());
		Assert.assertEquals("L6:N7", sheet1.getCell("M17").getValue());
		Assert.assertEquals("L16:L17", sheet1.getMergedRegion("L16").getReferenceString());
		Assert.assertEquals("M16:N16", sheet1.getMergedRegion("M16").getReferenceString());
		Assert.assertEquals("M17:N19", sheet1.getMergedRegion("N18").getReferenceString());
		Assert.assertEquals(3, sheet1.getOverlapsMergedRegions(new CellRegion("K15:N19"), false).size());
		
	}
	
	@Test 
	public void testCopySkipBlank(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue("A");
		sheet1.getCell("A3").setValue("");
		sheet1.getCell("A4").setValue("D");
		
		sheet1.getCell("B1").setValue("E");
		sheet1.getCell("B2").setValue("F");
		sheet1.getCell("B3").setValue("G");
		sheet1.getCell("B4").setValue("H");
		sheet1.getCell("C1").setValue("E");
		sheet1.getCell("C2").setValue("F");
		sheet1.getCell("C3").setValue("G");
		sheet1.getCell("C4").setValue("H");
		PasteOption opt = new PasteOption();
		
		sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("B1"), opt);
		
		Assert.assertEquals("A", sheet1.getCell("B1").getValue());
		Assert.assertEquals(null, sheet1.getCell("B2").getValue());
		Assert.assertEquals("", sheet1.getCell("B3").getValue());
		Assert.assertEquals("D", sheet1.getCell("B4").getValue());
		
		opt.setSkipBlank(true);
		sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("C1"), opt);
		Assert.assertEquals("A", sheet1.getCell("C1").getValue());
		Assert.assertEquals("F", sheet1.getCell("C2").getValue());
		Assert.assertEquals("", sheet1.getCell("C3").getValue());
		Assert.assertEquals("D", sheet1.getCell("C4").getValue());
	}
	
	@Test 
	public void testCut(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue("A");
		sheet1.getCell("A2").setValue("B");
		sheet1.getCell("A3").setValue("C");
		sheet1.getCell("A4").setValue("D");
		
		sheet1.getCell("B1").setValue("E");
		sheet1.getCell("B2").setValue("F");
		sheet1.getCell("B3").setValue("G");
		sheet1.getCell("B4").setValue("H");
		
		PasteOption opt = new PasteOption();
		opt.setCut(true);
		sheet1.pasteCell(new SheetRegion(sheet1,"A1:A4"), new CellRegion("B1"), opt);
		
		Assert.assertEquals(null, sheet1.getCell("A1").getValue());
		Assert.assertEquals(null, sheet1.getCell("A2").getValue());
		Assert.assertEquals(null, sheet1.getCell("A3").getValue());
		Assert.assertEquals(null, sheet1.getCell("A4").getValue());
		
		Assert.assertEquals("A", sheet1.getCell("B1").getValue());
		Assert.assertEquals("B", sheet1.getCell("B2").getValue());
		Assert.assertEquals("C", sheet1.getCell("B3").getValue());
		Assert.assertEquals("D", sheet1.getCell("B4").getValue());
		
		sheet1.pasteCell(new SheetRegion(sheet1,"B1:B4"), new CellRegion("B2"), opt);
		Assert.assertEquals(null, sheet1.getCell("B1").getValue());
		Assert.assertEquals("A", sheet1.getCell("B2").getValue());
		Assert.assertEquals("B", sheet1.getCell("B3").getValue());
		Assert.assertEquals("C", sheet1.getCell("B4").getValue());
		Assert.assertEquals("D", sheet1.getCell("B5").getValue());
		
		sheet1.pasteCell(new SheetRegion(sheet1,"B2:B4"), new CellRegion("B1"), opt);
		Assert.assertEquals("A", sheet1.getCell("B1").getValue());
		Assert.assertEquals("B", sheet1.getCell("B2").getValue());
		Assert.assertEquals("C", sheet1.getCell("B3").getValue());
		Assert.assertEquals(null, sheet1.getCell("B4").getValue());
		Assert.assertEquals("D", sheet1.getCell("B5").getValue());
	}
	

}
