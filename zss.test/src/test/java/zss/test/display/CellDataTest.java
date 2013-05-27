package zss.test.display;

import java.util.Arrays;
import java.util.List;

import org.hamcrest.CoreMatchers;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ErrorCollector;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheet "cell-data".
 * 
 * @author Hawk
 *
 */
@RunWith(Parameterized.class)
public class CellDataTest extends DisplayExcelTest{
	
	@Rule
    public ErrorCollector collector = new ErrorCollector();
	
	public CellDataTest(String page){
		super(page);
		sheet = zss.as(Spreadsheet.class).getBook().getSheetAt(3);
	}

	@Parameters
	public static List<Object[]> data() {
		Object[][] data = new Object[][] { { "/display.zul" }, { "/display2003.zul"}};
		return Arrays.asList(data);
	}
	
	@Test
	public void testCellFormat(){
		
		String expected = null;
		//TODO revise to a loop style
		try{
			//number
			expected = "1,234.56";
			collector.checkThat(Ranges.range(sheet,1,1).getCellData().getFormatText(), CoreMatchers.equalTo(expected));
			
			//currency
			expected = "NT$1,234.56";
			collector.checkThat(Ranges.range(sheet,1,2).getCellData().getFormatText(), CoreMatchers.equalTo(expected));

			expected = "¥1,234.00";
			collector.checkThat(Ranges.range(sheet,1,3).getCellData().getFormatText(), CoreMatchers.equalTo(expected));

			expected = "2013/4/12";
			collector.checkThat(Ranges.range(sheet,1,4).getCellData().getFormatText(), CoreMatchers.equalTo(expected));

			expected = "6:12 下午";
			collector.checkThat(Ranges.range(sheet,1,5).getCellData().getFormatText(), CoreMatchers.equalTo(expected));

			expected = "12.3%";
			collector.checkThat(Ranges.range(sheet,1,6).getCellData().getFormatText(), CoreMatchers.equalTo(expected));

			expected = " 12/25";
			collector.checkThat(Ranges.range(sheet,3,1).getCellData().getFormatText(), CoreMatchers.equalTo(expected));

			expected = "1.00E+09";
			collector.checkThat(Ranges.range(sheet,3,2).getCellData().getFormatText(), CoreMatchers.equalTo(expected));

			expected = "2013.4.12";

			collector.checkThat(Ranges.range(sheet,3,3).getCellData().getFormatText(), CoreMatchers.equalTo(expected));

			expected = "(07) 350-4450";
			collector.checkThat(Ranges.range(sheet,3,4).getCellData().getFormatText(), CoreMatchers.equalTo(expected));
		}catch (Exception e){
			collector.checkThat(e.toString(), CoreMatchers.equalTo(expected));
		}
	}
	
	@Test
	public void testNamedRange(){
		Sheet sheet = zss.as(Spreadsheet.class).getBook().getSheetAt(3);
		
		collector.checkThat(Ranges.range(sheet,10,6).getCellData().getFormatText(), CoreMatchers.equalTo("10"));
		Range rangeMerged = Ranges.range(sheet, "RangeMerged");
		
		collector.checkThat(rangeMerged.getRow(), CoreMatchers.equalTo(11));
		collector.checkThat(rangeMerged.getCellEditText(), CoreMatchers.equalTo("1"));

		collector.checkThat(Ranges.range(sheet,10,2).getCellData().getFormatText(), CoreMatchers.equalTo("21"));
		collector.checkThat(Ranges.range(sheet, "TestRange1"), CoreMatchers.notNullValue());
		collector.checkThat(Ranges.range(sheet, "TestRange1").getRow(), CoreMatchers.equalTo(11));
		
	}
}
