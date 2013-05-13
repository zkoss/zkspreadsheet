package zss.test.display;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;

import zss.test.SpreadsheetAgent;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheet "cell-data".
 * 
 * @author Hawk
 *
 */
@RunWith(Parameterized.class)
public class CellDataTest extends DisplayExcelTest{

	public CellDataTest(String page){
		super(page);
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		ssAgent.selectSheet("cell-data");
		sheet = zss.as(Spreadsheet.class).getXBook().getWorksheetAt(3);
	}

	@Parameters
	public static List<Object[]> data() {
		Object[][] data = new Object[][] { { "/display.zul" }, { "/display2003.zul"}};
		return Arrays.asList(data);
	}
	
	/*
	 * TODO what is the difference among getText(), getFormatText(), getRichEditText(), getEditText()?
	 */
	@Test
	public void testCellFormat(){
		/*
		//number
		assertEquals("1,234.56" ,Ranges.range(sheet,1,1).getText().getString());
		//currency
		assertEquals("NT$1,234.56" ,Ranges.range(sheet,1,2).getText().getString());

		//getFormatText() throw nullpointerexception
		assertEquals("¥1,234.00" ,Ranges.range(sheet,1,3).getText().getString());
		
		assertEquals("2013/4/12" ,Ranges.range(sheet,1,4).getText().getString());
		
		assertEquals("6:12 下午" ,Ranges.range(sheet,1,5).getText().getString());
		
		assertEquals("12.3%" ,Ranges.range(sheet,1,6).getText().getString());
		
		
		assertEquals("12/25" ,Ranges.range(sheet,3,1).getText().getString());
		
		assertEquals("1.00E+09" ,Ranges.range(sheet,3,2).getText().getString());
		
		assertEquals("2013.4.12" ,Ranges.range(sheet,3,3).getText().getString());
		
		assertEquals("(07)350-4450" ,Ranges.range(sheet,3,4).getText().getString());
		*/
	}
	
	@Test
	public void testNamedRange(){
		Sheet sheet = zss.as(Spreadsheet.class).getBook().getSheetAt(3);
		
		assertEquals("10",Ranges.range(sheet,10,6).getCellValue().toString());
		Range rangeMerged = Ranges.range(sheet, "RangeMerged");
		
		assertEquals(11,rangeMerged.getRow());
		assertEquals("1",rangeMerged.getCellEditText());

		assertEquals("21",Ranges.range(sheet,10,2).getCellValue().toString());
		assertNotNull(Ranges.range(sheet, "TestRange1"));
		assertEquals(11, Ranges.range(sheet, "TestRange1").getRow());
		
	}
}
