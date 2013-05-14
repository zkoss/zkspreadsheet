package zss.test.display;

import static org.junit.Assert.assertEquals;

import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.zss.ui.Spreadsheet;

import zss.test.SpreadsheetAgent;


/**
 * Test case for the function "display Excel files" for 2007 and 2003
 * Testing for the xsheet "cell-border".
 * 
 * Because the specification of color is unclear, we don't test color-related feature.
 * 
 * @author Hawk
 *
 */
@RunWith(Parameterized.class)
public class CellBorderTest extends DisplayExcelTest{

	public CellBorderTest(String testPage){
		super(testPage);
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		ssAgent.selectSheet("cell-border");
		xsheet = zss.as(Spreadsheet.class).getXBook().getWorksheetAt(1);
	}

	@Parameters
	public static List<Object[]> data() {
		Object[][] data = new Object[][] { { "/display.zul" }, { "/display2003.zul"}};
		return Arrays.asList(data);
	}

	/*
	@Test
	public void testBorderColor(){
		
	}
	*/
	
	@Test
	public void testBorderPosition(){
		//bottom
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 2, 1).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 2, 1).getCellStyle().getBorderTop());

		//top
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 2, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 2, 2).getCellStyle().getBorderRight());
		
		//left
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 2, 3).getCellStyle().getBorderLeft());
		
		//right
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 2, 4).getCellStyle().getBorderRight());
	}
	
	@Test
	public void testBorderStyle(){
		//only styles that can be rendered 
		assertEquals(CellStyle.BORDER_HAIR,getCell(xsheet, 4, 1).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_DOTTED,getCell(xsheet, 4, 2).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_DASHED,getCell(xsheet, 4, 3).getCellStyle().getBorderBottom());
		
	}
	
	@Test
	public void testAllBorders4Cells(){
		for (int row = 8 ; row<=9 ; row++){
			for (int column=1 ; column <=2 ; column++){
				assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, row, column).getCellStyle().getBorderTop());
				assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, row, column).getCellStyle().getBorderLeft());
				assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, row, column).getCellStyle().getBorderRight());
				assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, row, column).getCellStyle().getBorderBottom());
			}
		}
	}
	
	@Test
	public void testOutlineBorders4Cells(){
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 8, 4).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 8, 4).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 8, 4).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 8, 4).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 8, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 8, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 8, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 8, 5).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 9, 4).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 9, 4).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 9, 4).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 9, 4).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 9, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 9, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 9, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 9, 5).getCellStyle().getBorderBottom());;
	}
	
	@Test
	public void testInsideBorders4Cells(){
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 11, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 11, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 11, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 11, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 11, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 11, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 11, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 11, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 12, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 12, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 12, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 12, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 12, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 12, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 12, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 12, 2).getCellStyle().getBorderBottom());;
	}
	
	@Test
	public void testOutlineBorders9Cells(){

		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 3).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 3).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 1).getCellStyle().getBorderRight());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 3).getCellStyle().getBorderBottom());
		
	}
	
	@Test
	public void testInsideBorders9Cells(){
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 5).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 6).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 6).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 6).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 6).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 7).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 7).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 18, 7).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 18, 7).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 5).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 6).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 6).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 6).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 6).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 7).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 7).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 19, 7).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 19, 7).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 5).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 6).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 6).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 6).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 6).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 7).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 20, 7).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 7).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 20, 7).getCellStyle().getBorderBottom());
	}
	
	@Test
	public void testHorizontalBorders9Cells(){
		for (int row=22; row<=24 ; row++){
			for (int column = 1 ; column<=3 ; column++){
				assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, row, column).getCellStyle().getBorderTop());
				assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, row, column).getCellStyle().getBorderLeft());
				assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, row, column).getCellStyle().getBorderRight());
				assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, row, column).getCellStyle().getBorderBottom());
			}
		}
	}
	
	@Test
	public void testVerticalBorders9Cells(){
		for (int row=22; row<=24 ; row++){
			for (int column = 5 ; column<=7 ; column++){
				assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, row, column).getCellStyle().getBorderTop());
				assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, row, column).getCellStyle().getBorderLeft());
				assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, row, column).getCellStyle().getBorderRight());
				assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, row, column).getCellStyle().getBorderBottom());
			}
		}
	}
	
	@Test
	public void testMergedCellBorder(){
		//merge 2 cells
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 27, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 1).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 27, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 2).getCellStyle().getBorderBottom());

		//merge 3 cells
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 4).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 4).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 27, 4).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 4).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 5).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 27, 5).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 27, 5).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 5).getCellStyle().getBorderBottom());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 6).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 27, 6).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 6).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 27, 6).getCellStyle().getBorderBottom());
		
	}
	
	@Test
	public void testMergedVerticalCellBorder(){
		//merge cells vertically
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 30, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 30, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 30, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 30, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 31, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 31, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 31, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 31, 1).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 32, 1).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 32, 1).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 32, 1).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 32, 1).getCellStyle().getBorderBottom());

	}
	
	@Test
	public void testMergedVerticalHorizontalCellBorder(){
		//merge cells vertically and horizontally
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 30, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 30, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 30, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 30, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 30, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 30, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 30, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 30, 3).getCellStyle().getBorderBottom());
		
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 31, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 31, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 31, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 31, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 31, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 31, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 31, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 31, 3).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 32, 2).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 32, 2).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 32, 2).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 32, 2).getCellStyle().getBorderBottom());
		
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 32, 3).getCellStyle().getBorderTop());
		assertEquals(CellStyle.BORDER_NONE,getCell(xsheet, 32, 3).getCellStyle().getBorderLeft());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 32, 3).getCellStyle().getBorderRight());
		assertEquals(CellStyle.BORDER_THIN,getCell(xsheet, 32, 3).getCellStyle().getBorderBottom());

	}
}
