package zss.test.display;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;

import zss.test.SpreadsheetAgent;
import zss.test.SpreadsheetTestCaseBase;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheet "cell-data".
 * 
 * @author Hawk
 *
 */
public class CellDataTest extends SpreadsheetTestCaseBase{

	private static DesktopAgent desktop; 
	private static ComponentAgent zss ;
	private static Worksheet sheet;
	
	@BeforeClass
	public static void initialize(){
		//sheet 0
		desktop = Zats.newClient().connect("/display.zul");
		
		zss = desktop.query("spreadsheet");
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		ssAgent.selectSheet("cell-data");
		sheet = zss.as(Spreadsheet.class).getSheet(3);
	}
	

	/*
	 * TODO what is the difference among getText(), getFormatText(), getRichEditText(), getEditText()?
	 */
	@Test
	public void testCellFormat(){
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
	}
	
	@Test
	public void testNamedRange(){
		assertEquals("10",Ranges.range(sheet,10,6).getText().getString());
		Range rangeMerged = Ranges.range(sheet, "RangeMerged");
		assertEquals(11,rangeMerged.getRow());
		assertEquals("1",rangeMerged.getEditText());

		assertEquals("21",Ranges.range(sheet,10,2).getText().getString());
		assertNotNull(Ranges.range(sheet, "TestRange1"));
		assertEquals(11, Ranges.range(sheet, "TestRange1").getRow());
		
	}
}
