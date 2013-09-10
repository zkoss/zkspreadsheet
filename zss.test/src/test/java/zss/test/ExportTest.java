package zss.test;

import org.hamcrest.CoreMatchers;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ErrorCollector;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * For testing export.
 */
public class ExportTest extends SpreadsheetTestCaseBase{

	private DesktopAgent desktop;
	
	@Rule
    public ErrorCollector collector = new ErrorCollector();
	
	public ExportTest(){

		desktop = Zats.newClient().connect("/export.zul");
		
	}
	
	@Test
	public void testExport() {

		//press buttons to input data in the source spreadsheet
		desktop.query("#inputText").click();
		desktop.query("#inputPicture").click();
		
		//export to destination spreadsheet
		desktop.query("#exportImport").click();
		
		//compare two component's data
		
		ComponentAgent srcZss = desktop.query("#source");
		ComponentAgent dstZss = desktop.query("#destination");
		
		Sheet srcSheet = srcZss.as(Spreadsheet.class).getSelectedSheet();
		Sheet dstSheet = dstZss.as(Spreadsheet.class).getSelectedSheet();
		for (int row = 0 ; row < 16 ; row ++){
			for (int col =0 ; col < 12 ; col++){
				Object expected = Ranges.range(srcSheet, row, col).getCellEditText();
				Object actual =  Ranges.range(dstSheet, row, col).getCellEditText();
				collector.checkThat("unequal cell "+row+","+col, actual, CoreMatchers.equalTo(expected));
				
				//check style
				expected = Ranges.range(srcSheet, row, col).getCellStyle();
				
				actual =  Ranges.range(dstSheet, row, col).getCellStyle();
				collector.checkThat("unequal cell "+row+","+col, actual, CoreMatchers.equalTo(expected));
			}
		}
		
		collector.checkThat(dstSheet.getPictures().size()
				, CoreMatchers.equalTo(srcSheet.getPictures().size()));
	}
	
}
