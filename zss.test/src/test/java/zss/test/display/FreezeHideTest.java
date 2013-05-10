package zss.test.display;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Spreadsheet;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheets "row", "column", "rowcolumn".
 * 
 * @author Hawk
 *
 */
@RunWith(Parameterized.class)
public class FreezeHideTest extends DisplayExcelTest{

	private Spreadsheet spreadsheet;
	
	public FreezeHideTest(String testPage){
		super(testPage);
		spreadsheet = zss.as(Spreadsheet.class);
	}

	@Parameters
	public static List<Object[]> data() {
		Object[][] data = new Object[][] { { "/display.zul" }, { "/display2003.zul"}};
		return Arrays.asList(data);
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
	
		assertEquals(true,isHiddenRow(sheet, 5));
	}
	
	
	
	@Test
	public void testHiddenColumn(){
		Worksheet sheet = spreadsheet.getBook().getWorksheet("column");
	
		assertFalse(isHiddenColumn(sheet, 0));
		//check the hidden column
		assertTrue(isHiddenColumn(sheet, 4));
	}
	
	@Test
	public void testHiddenRowColumn(){
		Worksheet sheet = spreadsheet.getBook().getWorksheet("rowcolumn");
	
		assertFalse(isHiddenRow(sheet, 0));
		assertTrue(isHiddenRow(sheet, 5));
		
		assertFalse(isHiddenColumn(sheet, 0));
		assertTrue(isHiddenColumn(sheet, 4));
	}
	
	private int getFrozenRow(Worksheet sheet) {
		return BookHelper.getRowFreeze(sheet);
	}

	private int getFrozenColumn(Worksheet sheet) {
		return BookHelper.getColumnFreeze(sheet);
	}
}
