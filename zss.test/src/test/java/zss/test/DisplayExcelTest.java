package zss.test;

import static org.junit.Assert.assertEquals;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;

import com.lowagie.text.Cell;

/**
 * Test case for the function "display Excel files".
 * Testing for the sheet "cell-text".
 * 
 * Because the specification of color is unclear, we don't test color-related feature.
 * 
 * @author Hawk
 *
 */
public class DisplayExcelTest extends SpreadsheetTestCaseBase{

	private static DesktopAgent desktop; 
	private static ComponentAgent zss ;
	private static Spreadsheet spreadsheet ;
	private static Worksheet sheet;
	private static Book book;
	
	@BeforeClass
	public static void initialize(){
		//sheet 0
		desktop = Zats.newClient().connect("/display.zul");
		
		zss = desktop.query("spreadsheet");
		spreadsheet = zss.as(Spreadsheet.class);
		sheet = zss.as(Spreadsheet.class).getSheet(0);
		book = spreadsheet.getBook();
		
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
		Font font = book.getFontAt(getCell(sheet, 3, 0).getCellStyle().getFontIndex());
		assertEquals("Arial", font.getFontName());
		font = book.getFontAt(getCell(sheet, 3, 1).getCellStyle().getFontIndex());
		assertEquals("Arial Black", font.getFontName());
		font = book.getFontAt(getCell(sheet, 3, 2).getCellStyle().getFontIndex());
		assertEquals("Calibri", font.getFontName());
		font = book.getFontAt(getCell(sheet, 3, 3).getCellStyle().getFontIndex());
		assertEquals("Comic Sans MS", font.getFontName());
		font = book.getFontAt(Utils.getCell(sheet, 3, 4).getCellStyle().getFontIndex());
		assertEquals("Courier New", font.getFontName());
		font = book.getFontAt(Utils.getCell(sheet, 3, 5).getCellStyle().getFontIndex());
		assertEquals("Georgia", font.getFontName());
		font = book.getFontAt(Utils.getCell(sheet, 3, 6).getCellStyle().getFontIndex());
		assertEquals("Impact", font.getFontName());
		font = book.getFontAt(Utils.getCell(sheet, 3, 7).getCellStyle().getFontIndex());
		assertEquals("Lucida Console", font.getFontName());
		font = book.getFontAt(Utils.getCell(sheet, 3, 8).getCellStyle().getFontIndex());
		assertEquals("Lucida Sans Unicode", font.getFontName());
		font = book.getFontAt(Utils.getCell(sheet, 3, 9).getCellStyle().getFontIndex());
		assertEquals("Palatino Linotype", font.getFontName());
		font = book.getFontAt(Utils.getCell(sheet, 3, 10).getCellStyle().getFontIndex());
		assertEquals("Tahoma", font.getFontName());
		
		font = book.getFontAt(Utils.getCell(sheet, 4, 0).getCellStyle().getFontIndex());
		assertEquals("Times New Roman", font.getFontName());
		font = book.getFontAt(Utils.getCell(sheet, 4, 1).getCellStyle().getFontIndex());
		assertEquals("Trebuchet MS", font.getFontName());
		font = book.getFontAt(Utils.getCell(sheet, 4, 2).getCellStyle().getFontIndex());
		assertEquals("Verdana", font.getFontName());
		font = book.getFontAt(Utils.getCell(sheet, 4, 3).getCellStyle().getFontIndex());
		assertEquals("Microsoft Sans Serif", font.getFontName());
//		font = book.getFontAt(Utils.getCell(sheet, 4, 4).getCellStyle().getFontIndex());
//		assertEquals("Ms Serif", font.getFontName()); non-existed font 
		font = book.getFontAt(Utils.getCell(sheet, 4, 5).getCellStyle().getFontIndex());
		assertEquals("Book Antiqua", font.getFontName());
	}

	@Test
	public void testFontSize(){
		Font font = book.getFontAt(getCell(sheet, 6, 0).getCellStyle().getFontIndex());
		assertEquals(8, font.getFontHeight()/20);
		
		font = book.getFontAt(getCell(sheet, 6, 1).getCellStyle().getFontIndex());
		assertEquals(12, font.getFontHeight()/20);
		
		font = book.getFontAt(getCell(sheet, 6, 2).getCellStyle().getFontIndex());
		assertEquals(28, font.getFontHeight()/20);
		
		font = book.getFontAt(getCell(sheet, 6, 3).getCellStyle().getFontIndex());
		assertEquals(72, font.getFontHeight()/20);
	}
	
	
	@Test
	public void testFontStyle(){
		Font font = book.getFontAt(getCell(sheet,9, 0).getCellStyle().getFontIndex());
		assertEquals(Font.BOLDWEIGHT_BOLD, font.getBoldweight());
		
		font = book.getFontAt(getCell(sheet,9, 1).getCellStyle().getFontIndex());
		assertEquals(true, font.getItalic());
		
		font = book.getFontAt(getCell(sheet,9, 2).getCellStyle().getFontIndex());
		assertEquals(true, font.getStrikeout());
		
		font = book.getFontAt(getCell(sheet,9, 3).getCellStyle().getFontIndex());
		assertEquals(Font.U_SINGLE, font.getUnderline());
		font = book.getFontAt(getCell(sheet,9, 4).getCellStyle().getFontIndex());
		assertEquals(Font.U_DOUBLE, font.getUnderline());
		font = book.getFontAt(getCell(sheet,9, 5).getCellStyle().getFontIndex());
		assertEquals(Font.U_SINGLE_ACCOUNTING, font.getUnderline());
		font = book.getFontAt(getCell(sheet,9, 6).getCellStyle().getFontIndex());
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
		Font font = book.getFontAt(getCell(sheet, 14, 1).getCellStyle().getFontIndex());
		assertEquals("Georgia", font.getFontName());
		font = book.getFontAt(getCell(sheet, 15, 1).getCellStyle().getFontIndex());
		assertEquals("Georgia", font.getFontName());
		
		font = book.getFontAt(getCell(sheet, 14, 2).getCellStyle().getFontIndex());
		assertEquals("Times New Roman", font.getFontName());
		font = book.getFontAt(getCell(sheet, 15, 2).getCellStyle().getFontIndex());
		assertEquals("Times New Roman", font.getFontName());
	}
	
	@Test
	public void testAlignment(){
		//horizontal alignment
		assertEquals(Cell.ALIGN_LEFT,getCell(sheet,27,1).getCellStyle().getAlignment());
	}
}
