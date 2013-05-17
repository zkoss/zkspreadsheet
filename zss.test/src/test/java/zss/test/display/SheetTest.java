package zss.test.display;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.XUtils;

import zss.test.SpreadsheetAgent;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheet "sheet-protection", "sheet-autofilter".
 * 
 * @author Hawk
 *
 */
@RunWith(Parameterized.class)
public class SheetTest extends DisplayExcelTest{

	public SheetTest(String testPage){
		super(testPage);
	}

	@Parameters
	public static List<Object[]> data() {
		Object[][] data = new Object[][] { { "/display.zul" }, { "/display2003.zul"}};
		return Arrays.asList(data);
	}
	

	@Test
	public void testProtection(){
		
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		ssAgent.selectSheet("sheet-protection");
		xsheet = zss.as(Spreadsheet.class).getXBook().getWorksheet("sheet-protection");

		assertTrue(xsheet.getProtect());
		assertTrue(getCell(xsheet, 0, 0).getCellStyle().getLocked());
		//check editable cell
		assertFalse(getCell(xsheet, 1, 0).getCellStyle().getLocked());
	}
	
	@Test
	public void testAutofilter(){
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		ssAgent.selectSheet("sheet-autofilter");
		xsheet = zss.as(Spreadsheet.class).getXBook().getWorksheet("sheet-autofilter");
		
		AutoFilter autoFilter = xsheet.getAutoFilter(); 
		assertNotNull(autoFilter);
		assertTrue(autoFilter.getFilterColumn(1).isOn());
		assertEquals(1, autoFilter.getFilterColumn(1).getFilters().size());
		assertEquals("Sunday", autoFilter.getFilterColumn(1).getFilters().get(0));
		for (int r1=1 ; r1<=6 ; r1++){
			assertTrue(isHiddenRow(xsheet, r1));
		}
		for (int r2=8 ; r2<=13 ; r2++){
			assertTrue(isHiddenRow(xsheet, r2));
		}
	}
	
	private Cell getCell(XSheet sheet, int row, int col) {
		return XUtils.getCell(sheet, row, col);
	}
}
