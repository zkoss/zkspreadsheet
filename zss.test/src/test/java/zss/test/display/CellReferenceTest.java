package zss.test.display;

import static org.junit.Assert.assertEquals;

import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheet "cell-reference".
 * 
 * @author Hawk
 *
 */
@RunWith(Parameterized.class)
public class CellReferenceTest extends DisplayExcelTest{

	public CellReferenceTest(String page){
		super(page);
	}
	
	@Parameters
	public static List<Object[]> getTestPages() {
		Object[][] data = new Object[][] { {"/external-reference.zul?year=2007" }, { "/external-reference.zul?year=2003"}};
		return Arrays.asList(data);
	}

	@Test
	public void testCrossSheetReference(){
		zss = desktop.query("spreadsheet");
		Sheet sheet = zss.as(Spreadsheet.class).getBook().getSheet("cell-reference");
		
		assertEquals("The first row is freezed.", Ranges.range(sheet, 3, 2).getCellData().getFormatText());
		assertEquals("9", Ranges.range(sheet, 4, 2).getCellData().getFormatText());
	}
	
	@Test
	public void testExternalFileReference(){
		zss = desktop.query("#dst");
		Sheet sheet = zss.as(Spreadsheet.class).getBook().getSheet("cell-reference");
		
		assertEquals("ZK", Ranges.range(sheet, 0, 2).getCellData().getFormatText());
		assertEquals("10", Ranges.range(sheet, 1, 2).getCellData().getFormatText());
	}
}
