package zss.test.display;

import static org.junit.Assert.assertEquals;

import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
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
		XSheet xsheet = spreadsheet.getXBook().getWorksheet("row");
		//the number of frozen row
		assertEquals(1, getFrozenRow(xsheet));
	
	}
	
	
	@Test
	public void testFrozenColumn(){
		XSheet xsheet = spreadsheet.getXBook().getWorksheet("column");
		
		assertEquals(1, getFrozenColumn(xsheet));
	}
	
	@Test
	public void testFrozenRowColumn(){
		XSheet xsheet = spreadsheet.getXBook().getWorksheet("rowcolumn");
		
		assertEquals(2, getFrozenRow(xsheet));
		assertEquals(2, getFrozenColumn(xsheet));
	}
	
	//http://books.zkoss.org/wiki/ZK_Spreadsheet_Essentials/Working_with_ZK_Spreadsheet/Configure_and_Control_ZK_Spreadsheet/Hide_Row_and_Column_Titles
	@Test
	public void testHiddenRow(){
		Sheet sheet = spreadsheet.getBook().getSheet("row");
		assertEquals(true,sheet.isRowHidden(5));
	}
	
	
	
	@Test
	public void testHiddenColumn(){
		Sheet sheet = spreadsheet.getBook().getSheet("column");
		assertEquals(false,sheet.isColumnHidden(0));
		//check the hidden column
		assertEquals(true,sheet.isColumnHidden(4));
	}
	
	@Test
	public void testHiddenRowColumn(){
		Sheet sheet = spreadsheet.getBook().getSheet("rowcolumn");
		assertEquals(false,sheet.isRowHidden(0));
		assertEquals(true,sheet.isRowHidden(5));
		
		assertEquals(false,sheet.isColumnHidden(0));
		assertEquals(true,sheet.isColumnHidden(4));
	}
	
	private int getFrozenRow(XSheet xsheet) {
		return BookHelper.getRowFreeze(xsheet);
	}

	private int getFrozenColumn(XSheet xsheet) {
		return BookHelper.getColumnFreeze(xsheet);
	}
}
