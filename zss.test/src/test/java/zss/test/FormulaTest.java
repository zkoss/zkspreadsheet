package zss.test;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.junit.Assert.fail;

import org.hamcrest.number.IsCloseTo;
import org.junit.BeforeClass;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ErrorCollector;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
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
	
	private static final String END_OF_FORMULA="END OF FORMULA";
	private static final double DELTA = 1e-5;
	
	private static final int MAX_ROW_READ_FROM_LAST_FORUMULA = 30;
	
	@Rule
    public ErrorCollector collector = new ErrorCollector();
	
	@BeforeClass
	public static void initialize(){
		desktop = Zats.newClient().connect("/formula.zul");
		
		zss = desktop.query("spreadsheet");
		spreadsheet = zss.as(Spreadsheet.class);
		
	}
	
//	@Test
	public void testFirst() {
		Worksheet sheet = spreadsheet.getSheet(0);
		System.out.println("start test formulas in sheet: "+sheet.getSheetName());
//		int nFormula = testFormulaByRangesInSheet(sheet);
		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
//	@Test
	public void testTemp() {
		Worksheet sheet = spreadsheet.getSheet(1);
		System.out.println("start test formulas in sheet: "+sheet.getSheetName());
//		int nFormula = testFormulaByRangesInSheet(sheet);
		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testMath() {
		Worksheet sheet = spreadsheet.getSheet(1);
		
		int nFormula = testFormulaByRangesInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testLogical() {
		Worksheet sheet = spreadsheet.getSheet(2); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testText() {
		Worksheet sheet = spreadsheet.getSheet(3); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testInfo() {
		Worksheet sheet = spreadsheet.getSheet(4); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}

	@Test
	public void testDateTime() {
		Worksheet sheet = spreadsheet.getSheet(5); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testFinancial() {
		Worksheet sheet = spreadsheet.getSheet(6); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testStatistical() {
		Worksheet sheet = spreadsheet.getSheet(7); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}	

	@Test
	public void testEngineering() {
		Worksheet sheet = spreadsheet.getSheet(8); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}	
	
//	@Test
	public void testUnsupported() {
		Worksheet sheet = spreadsheet.getSheet(9); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}	
	


	//verify in text
	private int testFormulaByRangesInSheet(Worksheet sheet) {
		int nRowFromLastFormula = 0; //reset when read a formula cell
		int nFormula = 0;
		int row = 0 ; //current working row
		while (!isNoMoreFormula(nRowFromLastFormula)){
			Cell formulaCell = getCell(sheet, row, 1);
			Cell verifyCell = getCell(sheet, row, 3);
			if (isFormulaRow(verifyCell)
					&& formulaCell!=null && formulaCell.getCellType() == Cell.CELL_TYPE_FORMULA){ 
				String formulaResult = null;
				
				try{
					formulaResult = Ranges.range(sheet, row, 1).getText().getString();
				}catch (Exception e) {
					formulaResult = e.getMessage();
				}
				String expected = Ranges.range(sheet, row, 2).getText().getString();
				
				collector.checkThat(getFailedReason(formulaCell),expected, equalTo(formulaResult));
				nFormula ++;
				nRowFromLastFormula = 0;
			}
			row ++;
			nRowFromLastFormula++;
		}
		return nFormula;
	}	
	
	//verify each formula according cell's value
	private int testFormulaByPoiInSheet(Worksheet sheet) {
		int nRowFromLastFormula = 0; //reset when read a formula cell
		int nFormula = 0;
		String nameColumnText = null;
		int row = 0 ; //current working row
		while (!isNoMoreFormula(nRowFromLastFormula)){
			nameColumnText = Ranges.range(sheet, row, 0).getEditText();
			Cell formulaCell = getCell(sheet, row, 1);
			Cell expectedCell = getCell(sheet, row, 2);
			Cell verifyCell = getCell(sheet, row, 3);
			if (isFormulaRow(verifyCell)
					&& formulaCell!=null && formulaCell.getCellType() == Cell.CELL_TYPE_FORMULA){ 
				switch(formulaCell.getCachedFormulaResultType()){

				case Cell.CELL_TYPE_NUMERIC:
					collector.checkThat(getFailedReason(formulaCell),
							expectedCell.getNumericCellValue(), IsCloseTo.closeTo(formulaCell.getNumericCellValue(), DELTA));
					break;
				case Cell.CELL_TYPE_STRING:
					collector.checkThat(getFailedReason(formulaCell),
							expectedCell.getStringCellValue(), equalTo(formulaCell.getStringCellValue()));
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					collector.checkThat(getFailedReason(formulaCell),
							expectedCell.getBooleanCellValue(), equalTo(formulaCell.getBooleanCellValue()));
					break;
				case Cell.CELL_TYPE_ERROR:
					collector.checkThat(getFailedReason(formulaCell),
							Byte.toString(formulaCell.getErrorCellValue()), equalTo(getCellValue(expectedCell).toString()));
					break;
				default:
					fail("unsupported cell type:"+formulaCell.getCachedFormulaResultType()+" at row "+row+", column 1");
				}
				nFormula ++;
				nRowFromLastFormula = 0;				
			}
			row ++;
			nRowFromLastFormula++;
		}
		return nFormula;
	}
	
	private Object getCellValue(Cell cell) {
		switch(cell.getCellType()){
		case Cell.CELL_TYPE_NUMERIC:
			return cell.getNumericCellValue();
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		}
		return "TODO";
	}
	

	private boolean isFormulaRow(Cell verifyCell) {
		//skip "human check" cases
		return verifyCell != null  && verifyCell.getCellType() == Cell.CELL_TYPE_FORMULA ;
	}


	private String getFailedReason(Cell cell) {
		return cell.getCellFormula()+" failed, at row "+(cell.getRowIndex()+1)+" in sheet: "+cell.getSheet().getSheetName();
	}


	private boolean isNoMoreFormula(int nRowFromLastFormula) {
		return nRowFromLastFormula > MAX_ROW_READ_FROM_LAST_FORUMULA;
	}
}
