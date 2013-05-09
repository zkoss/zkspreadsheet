package zss.test;

import static org.hamcrest.CoreMatchers.equalTo;

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
	
	//tolerance delta value between correct value and expected value for a double number
	private static final double DELTA = 1e-5;
	
	// after reading the number of row without a formula, the program will stop to read more rows
	private static final int MAX_ROW_READ_FROM_LAST_FORUMULA = 30;
	
	@Rule
    public ErrorCollector collector = new ErrorCollector();
	
	@BeforeClass
	public static void initialize(){
		desktop = Zats.newClient().connect("/formula.zul");
		
		zss = desktop.query("spreadsheet");
		spreadsheet = zss.as(Spreadsheet.class);
		
	}
	

	@Test
	public void testMath() {
		Worksheet sheet = spreadsheet.getSheet(2);
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testLogical() {
		Worksheet sheet = spreadsheet.getSheet(3); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testText() {
		Worksheet sheet = spreadsheet.getSheet(4); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testInfo() {
		Worksheet sheet = spreadsheet.getSheet(5); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}

	@Test
	public void testDateTime() {
		Worksheet sheet = spreadsheet.getSheet(6); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testFinancial() {
		Worksheet sheet = spreadsheet.getSheet(7); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	@Test
	public void testStatistical() {
		Worksheet sheet = spreadsheet.getSheet(8); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}	

	@Test
	public void testEngineering() {
		Worksheet sheet = spreadsheet.getSheet(9); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}	
	
	@Test
	public void testUnsupported() {
		Worksheet sheet = spreadsheet.getSheet(10); 
		
		int nFormula = testFormulaByRangesInSheet(sheet);
//		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}	
	
	@Test
	public void testCustom() {
		Worksheet sheet = spreadsheet.getSheet(11); 
		
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
				
				collector.checkThat(getFailedReason(formulaCell),formulaResult, equalTo(expected));
				nFormula ++;
				nRowFromLastFormula = 0;
			}
			row ++;
			nRowFromLastFormula++;
		}
		return nFormula;
	}	
	
	
	/*
	 * Some formulas that cannot pass the test by Ranges API will also failed the test by POI API only when those formulas
	 * are put in the "first" sheet.  If they are not in the "first" sheet, the error won't occur.
	 * The root cause is that those specific formulas cells' getCachedFormulaResultType() always return 
	 * CELL_TYPE_ERROR in the first sheet, but when those formulas are not put in first sheet, 
	 * getCachedFormulaResultType() returns correct type.
	 * 
	 * Perform testFirst() and testTemp() can reproduce the error mentioned above.
	 */
//	@Test
	public void testFirst() {
		Worksheet sheet = spreadsheet.getSheet(0);
		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
//	@Test
	public void testTemp() {
		Worksheet sheet = spreadsheet.getSheet(1);
		int nFormula = testFormulaByPoiInSheet(sheet);
		System.out.println("tested "+nFormula+ " formulas in sheet: "+sheet.getSheetName());
	}
	
	
	
	//verify each formula according cell's value
	private int testFormulaByPoiInSheet(Worksheet sheet) {
		int nRowFromLastFormula = 0; //reset when read a formula cell
		int nFormula = 0;
		int row = 0 ; //current working row
		while (!isNoMoreFormula(nRowFromLastFormula)){
			Cell formulaCell = getCell(sheet, row, 1);
			Cell expectedCell = getCell(sheet, row, 2);
			Cell verifyCell = getCell(sheet, row, 3);
			if (isFormulaRow(verifyCell)
					&& formulaCell!=null && formulaCell.getCellType() == Cell.CELL_TYPE_FORMULA){ 
				if (formulaCell.getCachedFormulaResultType() == Cell.CELL_TYPE_NUMERIC){
					collector.checkThat(getFailedReason(formulaCell),formulaCell.getNumericCellValue()
							, IsCloseTo.closeTo(expectedCell.getNumericCellValue(),DELTA));
				}else{
					collector.checkThat(getFailedReason(formulaCell),getCellValue(formulaCell), equalTo(getCellValue(expectedCell)));
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
		int cellType = cell.getCellType() == Cell.CELL_TYPE_FORMULA ? cell.getCachedFormulaResultType():cell.getCellType();
		switch(cellType){
		case Cell.CELL_TYPE_NUMERIC:
			return cell.getNumericCellValue();
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_BOOLEAN:
			return cell.getBooleanCellValue();
		case Cell.CELL_TYPE_ERROR:
			return cell.getErrorCellValue();
		default:
			return "";
		}
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
