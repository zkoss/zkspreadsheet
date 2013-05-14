package zss.test.display;

import java.util.Arrays;
import java.util.List;

import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
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
	
	/*
	
	//http://books.zkoss.org/wiki/ZK_Spreadsheet_Essentials/Working_with_ZK_Spreadsheet/Configure_and_Control_ZK_Spreadsheet/Freeze_Rows_and_Columns
	@Test
	public void testFrozenRow(){
		Worksheet xsheet = spreadsheet.getBook().getWorksheet("row");
		//the number of frozen row
		assertEquals(1, getFrozenRow(xsheet));
	
	}
	
	
	@Test
	public void testFrozenColumn(){
		Worksheet xsheet = spreadsheet.getBook().getWorksheet("column");
		
		assertEquals(1, getFrozenColumn(xsheet));
	}
	
	@Test
	public void testFrozenRowColumn(){
		Worksheet xsheet = spreadsheet.getBook().getWorksheet("rowcolumn");
		
		assertEquals(2, getFrozenRow(xsheet));
		assertEquals(2, getFrozenColumn(xsheet));
	}
	
	//http://books.zkoss.org/wiki/ZK_Spreadsheet_Essentials/Working_with_ZK_Spreadsheet/Configure_and_Control_ZK_Spreadsheet/Hide_Row_and_Column_Titles
	@Test
	public void testHiddenRow(){
		Worksheet xsheet = spreadsheet.getBook().getWorksheet("row");
	
		assertEquals(true,isHiddenRow(xsheet, 5));
	}
	
	
	
	@Test
	public void testHiddenColumn(){
		Worksheet xsheet = spreadsheet.getBook().getWorksheet("column");
	
		assertFalse(isHiddenColumn(xsheet, 0));
		//check the hidden column
		assertTrue(isHiddenColumn(xsheet, 4));
	}
	
	@Test
	public void testHiddenRowColumn(){
		Worksheet xsheet = spreadsheet.getBook().getWorksheet("rowcolumn");
	
		assertFalse(isHiddenRow(xsheet, 0));
		assertTrue(isHiddenRow(xsheet, 5));
		
		assertFalse(isHiddenColumn(xsheet, 0));
		assertTrue(isHiddenColumn(xsheet, 4));
	}
	
	private int getFrozenRow(Worksheet xsheet) {
		return BookHelper.getRowFreeze(xsheet);
	}

	private int getFrozenColumn(Worksheet xsheet) {
		return BookHelper.getColumnFreeze(xsheet);
	}
	*/
}
