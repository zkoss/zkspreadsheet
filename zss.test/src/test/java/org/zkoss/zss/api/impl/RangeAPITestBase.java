package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.Range.AutoFillType;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Hyperlink;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;

public class RangeAPITestBase {
	
	protected void testDeleteAllSheet(Book workbook) throws IOException {
		List<Sheet> sheetList = new ArrayList<Sheet>();
		for(int i = 0; i < workbook.getNumberOfSheets(); i++) {
			sheetList.add(workbook.getSheetAt(i));
		}
		for(Sheet sheet : sheetList) {
			Ranges.range(sheet).deleteSheet();
		}
	}
	
	protected void testDeleteSheet(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Ranges.range(sheet).deleteSheet();
	}
	
	protected void testCreateSheet(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Ranges.range(sheet).createSheet("New Sheet");
		assertTrue(workbook.getSheet("New Sheet") != null);
	}
	
	protected void testToCellRange(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rA1C3 = Ranges.range(sheet, "A1:C3");
		Range shiftedRange = rA1C3.toShiftedRange(10, 5);
		assertEquals(3, shiftedRange.getRowCount());
		assertEquals(3, shiftedRange.getColumnCount());
		assertEquals(10, shiftedRange.getRow());
		assertEquals(12, shiftedRange.getLastRow());
		assertEquals(5, shiftedRange.getColumn());
		assertEquals(7, shiftedRange.getLastColumn());
		
		Range cellRange = shiftedRange.toCellRange(2, 2);
		
		assertEquals(1, cellRange.getRowCount());
		assertEquals(1, cellRange.getColumnCount());
		assertEquals(12, cellRange.getRow());
		assertEquals(7, cellRange.getColumn());
	}
	
	protected void testToShiftRanged(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rA1C3 = Ranges.range(sheet, "A1:C3");
		Range shiftedRange = rA1C3.toShiftedRange(10, 5);
		assertEquals(10, shiftedRange.getRow());
		assertEquals(12, shiftedRange.getLastRow());
		assertEquals(5, shiftedRange.getColumn());
		assertEquals(7, shiftedRange.getLastColumn());
	}
	
	protected void testGetColumnCount(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertEquals(6, rangeA5F18.getColumnCount());
	}
	
	protected void testGetColumn(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertEquals(0, rangeA5F18.getColumn());
	}
	
	protected void testGetRow(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertEquals(4, rangeA5F18.getRow());
	}
	
	protected void testGetLastColumn(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertEquals(5, rangeA5F18.getLastColumn());
	}
	
	protected void testGetLastRow(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertEquals(17, rangeA5F18.getLastRow());
	}
	
	protected void testGetRowCount(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertEquals(14, rangeA5F18.getRowCount());
	}
	
	protected void testIsWholeRow(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");

		assertTrue(!rangeA5F18.isWholeRow());
		
		Range rowRange = rangeA5F18.toRowRange();
		assertTrue(rowRange.isWholeRow());
	}
	
	protected void testIsWholeColumn(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");

		assertTrue(!rangeA5F18.isWholeColumn());
		
		Range colRange = rangeA5F18.toColumnRange();
		assertTrue(colRange.isWholeColumn());
	}
	
	protected void testIsWholeSheet(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range sheetRange = Ranges.range(sheet);
		assertTrue(sheetRange.isWholeSheet());
	}
	
	protected void testClearContents(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA1 = Ranges.range(sheet, "A1");
		rangeA1.setCellEditText("ZK");
		assertEquals("ZK", rangeA1.getCellEditText());
		CellOperationUtil.applyFontColor(rangeA1, "#ff0000");
		CellOperationUtil.applyFontStrikeout(rangeA1, true);
		assertEquals("#ff0000", rangeA1.getCellStyle().getFont().getColor().getHtmlColor());
		assertTrue(rangeA1.getCellStyle().getFont().isStrikeout());
		rangeA1.clearContents();
		assertEquals(CellType.BLANK, rangeA1.getCellData().getType());
		assertEquals("#ff0000", rangeA1.getCellStyle().getFont().getColor().getHtmlColor());
		assertTrue(rangeA1.getCellStyle().getFont().isStrikeout());
	}
	
	protected void testClearStyle(Book workbook) throws IOException {
		
		Sheet sheet = workbook.getSheet("Sheet1");
		
		Range rangeA1 = Ranges.range(sheet, "A1");
		
		CellOperationUtil.applyFontColor(rangeA1, "#ff0000");
		CellOperationUtil.applyFontStrikeout(rangeA1, true);
		assertEquals("#ff0000", rangeA1.getCellStyle().getFont().getColor().getHtmlColor());
		assertTrue(rangeA1.getCellStyle().getFont().isStrikeout());
		
		rangeA1.clearStyles();
		assertEquals("#000000", rangeA1.getCellStyle().getFont().getColor().getHtmlColor());
		assertTrue(!rangeA1.getCellStyle().getFont().isStrikeout());
	}
	
	// FIXME
	protected void testHyperLink(Book workbook, String outFileName) throws IOException {

		Sheet sheet = workbook.getSheet("Sheet1");
		Range rA1 = Ranges.range(sheet, "A1");
		rA1.setCellHyperlink(HyperlinkType.URL, "http://www.zkoss.org", "ZK");
		
		Range rA2 = Ranges.range(sheet, "A2");
		// set "info@zkoss.org" will cause illegal argument exception
		// set "mailto:info@zkoss.org" is correct
		rA2.setCellHyperlink(HyperlinkType.EMAIL, "mailto:info@zkoss.org", "ZK");
		
		rA2 = Ranges.range(sheet, "A2");
		Hyperlink linkEmail = rA2.getCellHyperlink();
		assertTrue(linkEmail != null);
		assertEquals("mailto:info@zkoss.org", linkEmail.getAddress());
		assertEquals(HyperlinkType.EMAIL, linkEmail.getType());
		
		Range rA3 = Ranges.range(sheet, "A3");
		rA3.setCellHyperlink(HyperlinkType.FILE, "test.xlsx", "ZK");
		
		Range rA4 = Ranges.range(sheet, "A4");
		rA4.setCellHyperlink(HyperlinkType.DOCUMENT, "test.xlsx", "ZK");
		
		Util.export(workbook, outFileName);
		
		// reload
		workbook = Util.loadBook(outFileName);
		sheet = workbook.getSheet("Sheet1");
		
		rA1 = Ranges.range(sheet, "A1");
		Hyperlink linkURL = rA1.getCellHyperlink();
		assertTrue(linkURL != null);
		assertEquals("http://www.zkoss.org", linkURL.getAddress());
		assertEquals(HyperlinkType.URL, linkURL.getType());
		
		rA2 = Ranges.range(sheet, "A2");
		linkEmail = rA2.getCellHyperlink();
		assertTrue(linkEmail != null);
		assertEquals("mailto:info@zkoss.org", linkEmail.getAddress());
		assertEquals(HyperlinkType.EMAIL, linkEmail.getType());
		
		rA3 = Ranges.range(sheet, "A3");
		Hyperlink linkFile = rA3.getCellHyperlink();
		assertTrue(linkFile != null);
		assertEquals("test.xlsx", linkFile.getAddress());
		assertEquals(HyperlinkType.FILE, linkFile.getType());
		
		rA4 = Ranges.range(sheet, "A4");
		Hyperlink linkDoc = rA4.getCellHyperlink();
		assertTrue(linkDoc != null);
		assertEquals("test.xlsx", linkDoc.getAddress());
		assertEquals(HyperlinkType.DOCUMENT, linkDoc.getType());
		
		Util.export(workbook, outFileName);
		
		rA1.setCellEditText("URL");
		//assertTrue(rA1.getCellHyperlink() == null);
		assertEquals("URL", rA1.getCellEditText());
		
		rA2.setCellEditText("Email");
		//assertTrue(rA2.getCellHyperlink() == null);
		assertEquals("Email", rA2.getCellEditText());
		
		rA3.setCellEditText("File");
		//assertTrue(rA3.getCellHyperlink() == null);
		assertEquals("File", rA3.getCellEditText());
		
		rA4.setCellEditText("Document");
		//assertTrue(rA4.getCellHyperlink() == null);
		assertEquals("Document", rA4.getCellEditText());
		
		Util.export(workbook, outFileName);
		
	}
}
