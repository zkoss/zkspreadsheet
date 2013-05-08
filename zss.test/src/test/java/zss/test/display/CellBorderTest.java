package zss.test.display;

import static org.junit.Assert.assertEquals;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;

import zss.test.SpreadsheetTestCaseBase;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheet "cell-border".
 * 
 * Because the specification of color is unclear, we don't test color-related feature.
 * 
 * @author Hawk
 *
 */
public class CellBorderTest extends SpreadsheetTestCaseBase{

	private static DesktopAgent desktop; 
	private static ComponentAgent zss ;
	private static Worksheet sheet;
	
	@BeforeClass
	public static void initialize(){
		desktop = Zats.newClient().connect("/display.zul");
		
		zss = desktop.query("spreadsheet");
		sheet = zss.as(Spreadsheet.class).getSheet(1);
		
	}
	

//	@Test
//	public void testBorderColor(){
//		
//	}
	
	@Test
	public void testBorderPosition(){
		//bottom
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 2, 1).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 2, 1).getCellStyle().getBorderTop());

		//top
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 2, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 2, 2).getCellStyle().getBorderRight());
		
		//left
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 2, 3).getCellStyle().getBorderLeft());
		
		//right
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 2, 4).getCellStyle().getBorderRight());
	}
	
	@Test
	public void testBorderStyle(){
		//only styles that can be rendered 
		assertEquals(CellStyle.BORDER_HAIR,getCell(sheet, 4, 1).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_DOTTED,getCell(sheet, 4, 2).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_DASHED,getCell(sheet, 4, 3).getCellStyle().getBorderBottom());
		
	}
	
	@Test
	public void testAllBorders4Cells(){
		for (int row = 8 ; row<=9 ; row++){
			for (int column=1 ; column <=2 ; column++){
				assertEquals(CellStyle.BORDER_THIN,getCell(sheet, row, column).getCellStyle().getBorderTop());
				assertEquals(CellStyle.BORDER_THIN,getCell(sheet, row, column).getCellStyle().getBorderLeft());
				assertEquals(CellStyle.BORDER_THIN,getCell(sheet, row, column).getCellStyle().getBorderRight());
				assertEquals(CellStyle.BORDER_THIN,getCell(sheet, row, column).getCellStyle().getBorderBottom());
			}
		}
	}
	
	@Test
	public void testOutlineBorders4Cells(){
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 8, 4).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 8, 4).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 8, 4).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 8, 4).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 8, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 8, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 8, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 8, 5).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 9, 4).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 9, 4).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 9, 4).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 9, 4).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 9, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 9, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 9, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 9, 5).getCellStyle().getBorderBottom());;
	}
	
	@Test
	public void testInsideBorders4Cells(){
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 11, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 11, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 11, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 11, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 11, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 11, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 11, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 11, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 12, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 12, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 12, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 12, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 12, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 12, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 12, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 12, 2).getCellStyle().getBorderBottom());;
	}
	
	@Test
	public void testOutlineBorders9Cells(){

		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 3).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 3).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 1).getCellStyle().getBorderRight());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 3).getCellStyle().getBorderBottom());
		
	}
	
	@Test
	public void testInsideBorders9Cells(){
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 5).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 6).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 6).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 6).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 6).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 7).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 7).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 18, 7).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 18, 7).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 5).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 6).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 6).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 6).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 6).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 7).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 7).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 19, 7).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 19, 7).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 5).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 6).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 6).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 6).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 6).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 7).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 20, 7).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 7).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 20, 7).getCellStyle().getBorderBottom());
	}
	
	@Test
	public void testHorizontalBorders9Cells(){
		for (int row=22; row<=24 ; row++){
			for (int column = 1 ; column<=3 ; column++){
				assertEquals(CellStyle.BORDER_THIN,getCell(sheet, row, column).getCellStyle().getBorderTop());
				assertEquals(CellStyle.BORDER_NONE,getCell(sheet, row, column).getCellStyle().getBorderLeft());
				assertEquals(CellStyle.BORDER_NONE,getCell(sheet, row, column).getCellStyle().getBorderRight());
				assertEquals(CellStyle.BORDER_THIN,getCell(sheet, row, column).getCellStyle().getBorderBottom());
			}
		}
	}
	
	@Test
	public void testVerticalBorders9Cells(){
		for (int row=22; row<=24 ; row++){
			for (int column = 5 ; column<=7 ; column++){
				assertEquals(CellStyle.BORDER_NONE,getCell(sheet, row, column).getCellStyle().getBorderTop());
				assertEquals(CellStyle.BORDER_THIN,getCell(sheet, row, column).getCellStyle().getBorderLeft());
				assertEquals(CellStyle.BORDER_THIN,getCell(sheet, row, column).getCellStyle().getBorderRight());
				assertEquals(CellStyle.BORDER_NONE,getCell(sheet, row, column).getCellStyle().getBorderBottom());
			}
		}
	}
	
	@Test
	public void testMergedCellBorder(){
		//merge 2 cells
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 27, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 1).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 27, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 2).getCellStyle().getBorderBottom());

		//merge 3 cells
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 4).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 4).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 27, 4).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 4).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 27, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 27, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 5).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 6).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 27, 6).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 6).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 27, 6).getCellStyle().getBorderBottom());
		
	}
	
	@Test
	public void testMergedVerticalCellBorder(){
		//merge cells vertically
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 30, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 30, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 30, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 30, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 31, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 31, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 31, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 31, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 32, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 32, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 32, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 32, 1).getCellStyle().getBorderBottom());

	}
	
	@Test
	public void testMergedVerticalHorizontalCellBorder(){
		//merge cells vertically and horizontally
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 30, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 30, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 30, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 30, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 30, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 30, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 30, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 30, 3).getCellStyle().getBorderBottom());
		
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 31, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 31, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 31, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 31, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 31, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 31, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 31, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 31, 3).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 32, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 32, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 32, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 32, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 32, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(sheet, 32, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 32, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(sheet, 32, 3).getCellStyle().getBorderBottom());

	}
}
