package zss.test.display;

import static org.junit.Assert.assertEquals;

import java.util.*;

import org.junit.*;
import org.junit.runner.RunWith;
import org.junit.runners.*;
import org.junit.runners.Parameterized.Parameters;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.zss.model.*;
import org.zkoss.zss.ui.Spreadsheet;

import zss.test.SpreadsheetAgent;


/**
 * Test cases are organized by "Testcase Class per Feature" pattern.
 * Test case for the function "display Excel files".
 * Testing for the sheet "cell-text".
 * 
 * Because the specification of color is unclear, we don't test color-related feature.
 * 
 * @author Hawk
 *
 */

@RunWith(Parameterized.class)
public class CellTextTest extends DisplayExcelTest{

	public CellTextTest(String page){
		super(page);
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		ssAgent.selectSheet("cell-text");
		sheet = zss.as(Spreadsheet.class).getSheet(0);
	}

	@Parameters
	public static List<Object[]> data() {
		Object[][] data = new Object[][] { { "/display.zul" }, { "/display2003.zul"}};
		return Arrays.asList(data);
	}
	
	/*
	@Test
	public void testFontColor() {
		
		String text = null;
		Color fontColor = null;
		
		//red
		text = Ranges.range(sheet, 1, 0).getEditText();
		assertEquals("ff0000", text);
		fontColor = getCell(sheet, 1, 0).getCellStyle().getFillForegroundColorColor();
		
		XSSFColor cellFontColor = getCellFontColor(spreadsheet, 1, 0);
		assertEquals("FFFF0000", cellFontColor.getARGBHex());
		
		//green
		text = Ranges.range(sheet, 1, 1).getEditText();
		assertEquals("00ff00", text);
		
		cellFontColor = getCellFontColor(spreadsheet, 1, 1);
		assertEquals("FF00FF00", cellFontColor.getARGBHex());
		//blue
		text = Ranges.range(sheet, 1, 2).getEditText();
		assertEquals("0000ff", text);
		
		cellFontColor = getCellFontColor(spreadsheet, 1, 2);
		assertEquals("FF0000FF", cellFontColor.getARGBHex());
		

	}
	*/
	
	@Test
	public void testFontFamily(){
		Font font = getFont(sheet, 3, 0);
		assertEquals("Arial", font.getFontName());
		font = getFont(sheet, 3, 1);
		assertEquals("Arial Black", font.getFontName());
		font = getFont(sheet, 3, 2);
		assertEquals("Calibri", font.getFontName());
		font = getFont(sheet, 3, 3);
		assertEquals("Comic Sans MS", font.getFontName());
		font = getFont(sheet, 3, 4);
		assertEquals("Courier New", font.getFontName());
		font = getFont(sheet, 3, 5);
		assertEquals("Georgia", font.getFontName());
		font = getFont(sheet, 3, 6);
		assertEquals("Impact", font.getFontName());
		font = getFont(sheet, 3, 7);
		assertEquals("Lucida Console", font.getFontName());
		font = getFont(sheet, 3, 8);
		assertEquals("Lucida Sans Unicode", font.getFontName());
		font = getFont(sheet, 3, 9);
		assertEquals("Palatino Linotype", font.getFontName());
		font = getFont(sheet, 3, 10);
		assertEquals("Tahoma", font.getFontName());
		
		font = getFont(sheet, 4, 0);
		assertEquals("Times New Roman", font.getFontName());
		font = getFont(sheet, 4, 1);
		assertEquals("Trebuchet MS", font.getFontName());
		font = getFont(sheet, 4, 2);
		assertEquals("Verdana", font.getFontName());
		font = getFont(sheet, 4, 3);
		assertEquals("Microsoft Sans Serif", font.getFontName());
//		font = getFont(sheet, 4, 4).getCellStyle().getFontIndex());
//		assertEquals("Ms Serif", font.getFontName()); non-existed font 
		font = getFont(sheet, 4, 5);
		assertEquals("Book Antiqua", font.getFontName());
	}

	private Font getFont(Worksheet sheet, int row, int column) {
		return sheet.getBook().getFontAt(getCell(sheet, row, column).getCellStyle().getFontIndex());
	}

	@Ignore("require furthur confirmation")
	public void testFontSize(){
		Font font =getFont(sheet, 6, 0);
		assertEquals(8, font.getFontHeight()/20);
		
		font =getFont(sheet, 6, 1);
		assertEquals(12, font.getFontHeight()/20);
		
		font =getFont(sheet, 6, 2);
		assertEquals(28, font.getFontHeight()/20);
		
		font =getFont(sheet, 6, 3);
		assertEquals(72, font.getFontHeight()/20);
	}
	
	
	@Test
	public void testFontStyle(){
		Font font =getFont(sheet,9, 0);
		assertEquals(Font.BOLDWEIGHT_BOLD, font.getBoldweight());
		
		font =getFont(sheet,9, 1);
		assertEquals(true, font.getItalic());
		
		font =getFont(sheet,9, 2);
		assertEquals(true, font.getStrikeout());
		
		font =getFont(sheet,9, 3);
		assertEquals(Font.U_SINGLE, font.getUnderline());
		font =getFont(sheet,9, 4);
		assertEquals(Font.U_DOUBLE, font.getUnderline());
		font =getFont(sheet,9, 5);
		assertEquals(Font.U_SINGLE_ACCOUNTING, font.getUnderline());
		font =getFont(sheet,9, 6);
		assertEquals(Font.U_DOUBLE_ACCOUNTING, font.getUnderline());
	}
	
	/*
	@Test
	public void testCellBackgroundColor(){
//		assertEquals("", getCell(sheet, 11, 0).getCellStyle().getFillBackgroundColor());
		System.out.println(((XSSFColor)getCell(sheet, 11, 0).getCellStyle().getFillBackgroundColorColor()).getARGBHex());
		assertEquals("", book.getCellStyleAt(getCell(sheet, 11, 1).getCellStyle().getIndex()).getFillBackgroundColor());
	}
	*/
	
	@Test
	public void testMixedCellStyle(){
		Font font =getFont(sheet, 14, 1);
		assertEquals("Georgia", font.getFontName());
		font =getFont(sheet, 15, 1);
		assertEquals("Georgia", font.getFontName());
		
		font =getFont(sheet, 14, 2);
		assertEquals("Times New Roman", font.getFontName());
		font =getFont(sheet, 15, 2);
		assertEquals("Times New Roman", font.getFontName());
	}
	
	@Test
	public void testAlignment(){
		//vertical alignment
		assertEquals(CellStyle.VERTICAL_TOP, getCell(sheet,26,1).getCellStyle().getVerticalAlignment());
		assertEquals(CellStyle.VERTICAL_CENTER, getCell(sheet,26,2).getCellStyle().getVerticalAlignment());
		assertEquals(CellStyle.VERTICAL_BOTTOM, getCell(sheet,26,3).getCellStyle().getVerticalAlignment());

		//horizontal alignment
		assertEquals(CellStyle.ALIGN_LEFT,getCell(sheet,27,1).getCellStyle().getAlignment());
		assertEquals(CellStyle.ALIGN_CENTER,getCell(sheet,27,2).getCellStyle().getAlignment());
		assertEquals(CellStyle.ALIGN_RIGHT,getCell(sheet,27,3).getCellStyle().getAlignment());
		
		//combined alignment
		assertEquals(CellStyle.VERTICAL_TOP,getCell(sheet,28,1).getCellStyle().getVerticalAlignment());
		assertEquals(CellStyle.ALIGN_LEFT,getCell(sheet,28,1).getCellStyle().getAlignment());
		assertEquals(CellStyle.VERTICAL_CENTER,getCell(sheet,28,2).getCellStyle().getVerticalAlignment());
		assertEquals(CellStyle.ALIGN_CENTER,getCell(sheet,28,2).getCellStyle().getAlignment());
		assertEquals(CellStyle.VERTICAL_BOTTOM,getCell(sheet,28,3).getCellStyle().getVerticalAlignment());
		assertEquals(CellStyle.ALIGN_RIGHT,getCell(sheet,28,3).getCellStyle().getAlignment());
	}
	
	@Test
	public void testHyperlink(){
		assertEquals(Hyperlink.LINK_URL, Ranges.range(sheet, 30, 1).getHyperlink().getType());
		assertEquals("http://www.zkoss.org/", Ranges.range(sheet, 30, 1).getHyperlink().getAddress());
	}
}
