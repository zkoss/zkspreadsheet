package zss.test.display;

import static org.junit.Assert.assertEquals;

import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.zkoss.zss.ui.Spreadsheet;

import zss.test.SpreadsheetAgent;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheets "chart-image".
 * 
 * @author Hawk
 *
 */

public class ChartImageTest extends DisplayExcelTest{

	private static Spreadsheet spreadsheet;
	
	//you can change to /display2003.zul to run single test
	public ChartImageTest(){
		this("/display.zul");
	}

	protected ChartImageTest(String testPage){
		super(testPage);
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		//select the sheet first or the chart won't be initialized
		ssAgent.selectSheet("chart-image");
		spreadsheet = zss.as(Spreadsheet.class);
		xsheet = spreadsheet.getXBook().getWorksheet("chart-image");
		sheet = spreadsheet.getBook().getSheet("chart-image");
	}

	

	@Test
	public void testPicture(){
		
		assertEquals(1, sheet.getPictures().size());
//		assertEquals(1, xsheet.getPictures().size());
		// TODO wait for 3.0 API supported
//		assertEquals(0, xsheet.getPictures().get(0).getPreferredSize().getCol1());
//		assertEquals(1, xsheet.getPictures().get(0).getPreferredSize().getCol2());
//		assertEquals(1, xsheet.getPictures().get(0).getPreferredSize().getRow1());
//		assertEquals(4, xsheet.getPictures().get(0).getPreferredSize().getRow2());
		
	}
	
	@Test
	public void testCharts(){
		assertEquals(12, xsheet.getCharts().size());
	}
}
