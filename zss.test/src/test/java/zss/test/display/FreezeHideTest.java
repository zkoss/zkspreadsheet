package zss.test.display;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Spreadsheet;

import zss.test.SpreadsheetTestCaseBase;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheets "row", "column", "rowcolumn".
 * 
 * @author Hawk
 *
 */
public class FreezeHideTest extends SpreadsheetTestCaseBase{

	private static DesktopAgent desktop; 
	private static ComponentAgent zss ;
	private static Spreadsheet spreadsheet;
	
	@BeforeClass
	public static void initialize(){
		desktop = Zats.newClient().connect("/display.zul");
		zss = desktop.query("spreadsheet");
		spreadsheet = zss.as(Spreadsheet.class);
		
	}
	//http://books.zkoss.org/wiki/ZK_Spreadsheet_Essentials/Working_with_ZK_Spreadsheet/Configure_and_Control_ZK_Spreadsheet/Freeze_Rows_and_Columns
	@Test
	public void testFrozenRow(){
		Worksheet sheet = spreadsheet.getBook().getWorksheet("row");
		//the number of frozen row
		assertEquals(1, getFrozenRow(sheet));
	
	}
	
	
	@Test
	public void testFrozenColumn(){
		Worksheet sheet = spreadsheet.getBook().getWorksheet("column");
		
		assertEquals(1, getFrozenColumn(sheet));
	}
	
	@Test
	public void testFrozenRowColumn(){
		Worksheet sheet = spreadsheet.getBook().getWorksheet("rowcolumn");
		
		assertEquals(2, getFrozenRow(sheet));
		assertEquals(2, getFrozenColumn(sheet));
	}
	
	//http://books.zkoss.org/wiki/ZK_Spreadsheet_Essentials/Working_with_ZK_Spreadsheet/Configure_and_Control_ZK_Spreadsheet/Hide_Row_and_Column_Titles
	@Test
	public void testHiddenRow(){
		Worksheet sheet = spreadsheet.getBook().getWorksheet("row");
	
		assertEquals(true,isRowHidden(sheet, 5));
	}
	
	
	
	@Test
	public void testHiddenColumn(){
		Worksheet sheet = spreadsheet.getBook().getWorksheet("column");
	
		assertFalse(isColumnHidden(sheet, 0));
		//check the hidden column
		assertTrue(isColumnHidden(sheet, 4));
	}
	
	@Test
	public void testHiddenRowColumn(){
		Worksheet sheet = spreadsheet.getBook().getWorksheet("rowcolumn");
	
		assertFalse(isRowHidden(sheet, 0));
		assertTrue(isRowHidden(sheet, 5));
		
		assertFalse(isColumnHidden(sheet, 0));
		assertTrue(isColumnHidden(sheet, 4));
	}
	
	private boolean isColumnHidden(Worksheet sheet, int columnIndex) {
		return sheet.getColumnWidth(columnIndex)==0;
	}	
	
	private boolean isRowHidden(Worksheet sheet, int rowIndex) {
		return sheet.getRow(rowIndex).getZeroHeight();
	}	
	private int getFrozenRow(Worksheet sheet) {
		return BookHelper.getRowFreeze(sheet);
	}

	private int getFrozenColumn(Worksheet sheet) {
		return BookHelper.getColumnFreeze(sheet);
	}
}
