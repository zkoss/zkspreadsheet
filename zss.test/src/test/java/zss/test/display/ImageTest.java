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
@RunWith(Parameterized.class)
public class ImageTest extends DisplayExcelTest{

	private static Spreadsheet spreadsheet;
	
	public ImageTest(String testPage){
		super(testPage);
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		//select the sheet first or the chart won't be initialized
		ssAgent.selectSheet("chart-image");
		spreadsheet = zss.as(Spreadsheet.class);
		sheet = spreadsheet.getBook().getWorksheet("chart-image");
	}

	@Parameters
	public static List<Object[]> data() {
		Object[][] data = new Object[][] { { "/display.zul" }, { "/display2003.zul"}};
		return Arrays.asList(data);
	}
	

	@Test
	public void testPicture(){
		assertEquals(1, sheet.getPictures().size());
		assertEquals(0, sheet.getPictures().get(0).getPreferredSize().getCol1());
		assertEquals(1, sheet.getPictures().get(0).getPreferredSize().getCol2());
		assertEquals(1, sheet.getPictures().get(0).getPreferredSize().getRow1());
		assertEquals(4, sheet.getPictures().get(0).getPreferredSize().getRow2());
		
	}
	
	@Test
	public void testCharts(){
		assertEquals(12, sheet.getCharts().size());
	}
}
