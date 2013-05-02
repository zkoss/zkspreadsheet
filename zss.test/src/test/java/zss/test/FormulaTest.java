package zss.test;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.junit.Assert.assertEquals;

import org.junit.BeforeClass;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ErrorCollector;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zats.mimic.operation.AuAgent;
import org.zkoss.zats.mimic.operation.AuData;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * For testing formulas.
 * 
 * @author Hawk
 *
 */
public class FormulaTest extends SpreadsheetTestCaseBase{

	private static DesktopAgent desktop; 
	private static ComponentAgent zss ;
	private static Spreadsheet spreadsheet ;
	private static Worksheet sheet;
	private static Book book;
	
	private static final String END_OF_FORMULA="END OF FORMULA";
	private static final double DELTA = 1e-5;
	
	@Rule
    public ErrorCollector collector = new ErrorCollector();
	
	@BeforeClass
	public static void initialize(){
		desktop = Zats.newClient().connect("/formula.zul");
		
		zss = desktop.query("spreadsheet");
		spreadsheet = zss.as(Spreadsheet.class);
//		spreadsheet.setSelectedSheet("formula-math"); java.lang.IllegalStateException: publish() can be called only in an event listener
		sheet = spreadsheet.getSheet(1); //formula-math
		book = spreadsheet.getBook();
		
	}
	@Test
	public void testMath() {
		int nFormula = 0;
		String nameColumnText = null;
		int row = 0 ; //current working row
		while (!END_OF_FORMULA.equals(nameColumnText)){
			nameColumnText = Ranges.range(sheet, row, 0).getEditText();
			String formula = Ranges.range(sheet, row, 1).getEditText();
			String expected = Ranges.range(sheet, row, 2).getEditText();
			String verification = Ranges.range(sheet, row, 3).getEditText();
			if (nameColumnText != null && !nameColumnText.isEmpty() 
					&& formula!= null && formula.startsWith("=")
					&& !verification.contains("human")){ //skip "human check" cases
				Cell cell = getCell(sheet, row, 1);
				switch(cell.getCachedFormulaResultType()){

				case Cell.CELL_TYPE_NUMERIC:
					collector.checkThat(Double.parseDouble(expected), equalTo(cell.getNumericCellValue()));
					break;
				case Cell.CELL_TYPE_STRING:
					collector.checkThat(expected, equalTo(cell.getStringCellValue()));
					break;
				}
//				assertEquals(Double.parseDouble(expected), cell.getNumericCellValue(), DELTA);
				nFormula ++;
			}
			row ++;
		}
		System.out.println("tested "+nFormula+" formulas");
	}
}
