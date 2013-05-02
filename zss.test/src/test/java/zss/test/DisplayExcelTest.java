package zss.test;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.xssf.usermodel.XSSFColor;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;

/**
 * Test case for the function "display Excel files".
 * Testing for each sheet are separated in different methods.
 * 
 * @author Hawk
 *
 */
public class DisplayExcelTest extends SpreadsheetTestCaseBase{

	@Test
	public void testCellText() {
		//sheet 0
		DesktopAgent desktop = Zats.newClient().connect("/display.zul");

		ComponentAgent zss = desktop.query("spreadsheet");
		Spreadsheet spreadsheet = zss.as(Spreadsheet.class);
		Worksheet sheet = zss.as(Spreadsheet.class).getSelectedSheet();
		Book book = spreadsheet.getBook();
		
		
		//font color
		String text = null;
		Color fontColor = null;
		
		//red
		text = Ranges.range(sheet, 1, 0).getEditText();
		assertEquals("ff0000", text);
		fontColor = Utils.getCell(sheet, 1, 0).getCellStyle().getFillForegroundColorColor();
		
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
		
		
		
		//font family
		Font font = book.getFontAt(Utils.getCell(sheet, 3, 0).getCellStyle().getFontIndex());
		assertEquals("Arial", font.getFontName());
		
	}

	private XSSFColor getCellFontColor(Spreadsheet spreadsheet, int row, int column) {
		return (XSSFColor)BookHelper.getFontColor(spreadsheet.getBook(),
			spreadsheet.getBook().getFontAt(Utils.getCell(spreadsheet.getSelectedSheet(), row, column).getCellStyle().getFontIndex()));
	}
}
