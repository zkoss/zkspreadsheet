package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.zkoss.zss.api.CellOperationUtil.applyAlignment;
import static org.zkoss.zss.api.CellOperationUtil.applyBackgroundColor;
import static org.zkoss.zss.api.CellOperationUtil.applyBorder;
import static org.zkoss.zss.api.CellOperationUtil.applyFontBoldweight;
import static org.zkoss.zss.api.CellOperationUtil.applyFontColor;
import static org.zkoss.zss.api.CellOperationUtil.applyFontItalic;
import static org.zkoss.zss.api.CellOperationUtil.applyFontSize;
import static org.zkoss.zss.api.CellOperationUtil.applyFontStrikeout;
import static org.zkoss.zss.api.CellOperationUtil.applyFontUnderline;
import static org.zkoss.zss.api.CellOperationUtil.applyVerticalAlignment;
import static org.zkoss.zss.api.Ranges.range;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.CellStyle.BorderType;

/**
 * 
 * @author kuro
 *
 */
public class StyleTest {

	private static Book _workbook;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	private void loadBook(String filename) throws IOException {
		final InputStream is = ShiftTest.class.getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
	}
	
	private void export() throws IOException {
		Exporter excelExporter = Exporters.getExporter("excel");
		FileOutputStream fos = new FileOutputStream(new File(ShiftTest.class.getResource("").getPath() + "book/test.xlsx"));
		excelExporter.export(_workbook, fos);
	}
	
	@Ignore
	@Test
	public void testStyleAfterExport() throws IOException {
		final String filename = "book/blank.xlsx";
		final InputStream is =  getClass().getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
		
		Sheet sheet = _workbook.getSheet("Sheet1");
		
		Range rA1 = range(sheet, "A1");
		rA1.setCellEditText("Bold");
		applyFontBoldweight(rA1, Font.Boldweight.BOLD);
		
		Range rB1 = range(sheet, "B1");
		rB1.setCellEditText("Italic");
		applyFontItalic(rB1, true);
		
		Range rC1 = range(sheet, "C1");
		rC1.setCellEditText("Underline");
		applyFontUnderline(rC1, Font.Underline.SINGLE);
		
		Range rD1 = range(sheet, "D1");
		rD1.setCellEditText("Strikeout");
		applyFontStrikeout(rD1,  true);

		Range rE1 = range(sheet, "E1");
		rE1.setCellEditText("font #ff0000");
		applyFontColor(rE1, "#ff0000");
		rE1.setColumnWidth(100);
		
		Range rF1 = range(sheet, "F1");
		rF1.setCellEditText("background #00ff00");
		applyBackgroundColor(rF1, "#00ff00");
		rF1.setColumnWidth(150);

		Range rG1 = range(sheet, "G1");
		rG1.setCellEditText("center align");
		applyAlignment(rG1, CellStyle.Alignment.CENTER);
		applyVerticalAlignment(rG1, CellStyle.VerticalAlignment.TOP);
		rG1.setColumnWidth(100);
		rG1.setRowHeight(80);
		
		//input a table
		applyBorder(range(sheet, 2, 0, 6 ,3), ApplyBorderType.FULL, BorderType.THIN, "#000000");
		
		range(sheet, 2, 0).setCellEditText("Browser Market");
		range(sheet, 2,0 ,2,3).merge(false);
		applyFontSize(range(sheet, 2,0 ,2,3), (short)16);
		
		range(sheet, 3, 0).setCellEditText("Month");
		range(sheet, 3, 1).setCellEditText("IE");
		range(sheet, 3, 2).setCellEditText("Chrome");
		range(sheet, 3, 3).setCellEditText("Firefox");
		
		range(sheet, 4, 0).setCellEditText("Jan");
		range(sheet, 4, 1).setCellEditText("34");
		range(sheet, 4, 2).setCellEditText("26");
		range(sheet, 4, 3).setCellEditText("22");
		
		range(sheet, 5, 0).setCellEditText("Feb");
		range(sheet, 5, 1).setCellEditText("32");
		range(sheet, 5, 2).setCellEditText("27");
		range(sheet, 5, 3).setCellEditText("22");
		
		range(sheet, 6, 0).setCellEditText("Mar");
		range(sheet, 6, 1).setCellEditText("31");
		range(sheet, 6, 2).setCellEditText("28");
		range(sheet, 6, 3).setCellEditText("22");
		
		export();
		
		_workbook = null; 		// clean work book
		sheet = null;			// clean sheet
		loadBook("book/test.xlsx");  // Import
		sheet = _workbook.getSheet("Sheet1"); // get sheet	
		
		rA1 = range(sheet, "A1");
		assertEquals(Font.Boldweight.BOLD, rA1.getCellStyle().getFont().getBoldweight());
		
		rB1 = range(sheet, "B1");
		assertTrue(rB1.getCellStyle().getFont().isItalic());
		
		rC1 = range(sheet, "C1");
		assertEquals(Font.Underline.SINGLE, rC1.getCellStyle().getFont().getUnderline());
		
		rD1 = range(sheet, "D1");
		assertTrue(rD1.getCellStyle().getFont().isStrikeout());
		
		rE1 = range(sheet, "E1");
		assertEquals("#ff0000", rE1.getCellStyle().getFont().getColor().getHtmlColor());
		
		rF1 = range(sheet, "F1");
		assertEquals("#00ff00", rF1.getCellStyle().getBackgroundColor().getHtmlColor());
		
		rG1 = range(sheet, "G1");
		assertEquals(CellStyle.Alignment.CENTER, rG1.getCellStyle().getAlignment());
		assertEquals(CellStyle.VerticalAlignment.TOP, rG1.getCellStyle().getVerticalAlignment());
		
		// A3
		// assertEquals(BorderType.THIN, range(sheet, "A3").getCellStyle().getBorderLeft()); // should apply?
		assertEquals(BorderType.THIN, range(sheet, "A2").getCellStyle().getBorderBottom());
		// assertEquals(BorderType.THIN, range(sheet, "A3").getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, range(sheet, "A4").getCellStyle().getBorderTop());
		
		// B3
		assertEquals(BorderType.THIN, range(sheet, "B2").getCellStyle().getBorderBottom());
		// assertEquals(BorderType.THIN, range(sheet, "B3").getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, range(sheet, "B4").getCellStyle().getBorderTop());
		
		// C3
		assertEquals(BorderType.THIN, range(sheet, "C2").getCellStyle().getBorderBottom());
		// assertEquals(BorderType.THIN, range(sheet, "C3").getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, range(sheet, "C4").getCellStyle().getBorderTop());
	
		// D3
		assertEquals(BorderType.THIN, range(sheet, "E3").getCellStyle().getBorderLeft());
		assertEquals(BorderType.THIN, range(sheet, "D2").getCellStyle().getBorderBottom());
		// assertEquals(BorderType.THIN, range(sheet, "D3").getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, range(sheet, "D4").getCellStyle().getBorderTop());
		
		// Insert column C
		Range columnC = Ranges.range(sheet, "C1");
		columnC.toColumnRange().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		
		assertEquals(BorderType.THIN, range(sheet, "C2").getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, range(sheet, "C3").getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, range(sheet, "C4").getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, range(sheet, "C5").getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, range(sheet, "C6").getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, range(sheet, "C7").getCellStyle().getBorderBottom());
		
	}
}
