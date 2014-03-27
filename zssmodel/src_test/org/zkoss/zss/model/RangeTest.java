package org.zkoss.zss.model;

import static org.junit.Assert.*;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import java.util.concurrent.atomic.AtomicInteger;
import junit.framework.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.image.AImage;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.InvalidModelOpException;
import org.zkoss.zss.model.ModelEvent;
import org.zkoss.zss.model.ModelEventListener;
import org.zkoss.zss.model.ModelEvents;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SChart.ChartType;
import org.zkoss.zss.model.SHyperlink;
import org.zkoss.zss.model.SPicture;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.ViewAnchor;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.SHyperlink.HyperlinkType;
import org.zkoss.zss.model.SPicture.Format;
import org.zkoss.zss.model.chart.SGeneralChartData;
import org.zkoss.zss.model.chart.SSeries;
import org.zkoss.zss.range.*;
import org.zkoss.zss.range.SRange.DeleteShift;
import org.zkoss.zss.range.SRange.InsertCopyOrigin;
import org.zkoss.zss.range.SRange.InsertShift;

public class RangeTest {

	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}
	
	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	@Test
	public void testGetRange(){
		
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = book.createSheet("Sheet1");
		
		
		SRange r1 = SRanges.range(sheet1);
		Assert.assertEquals(0, r1.getRow());
		Assert.assertEquals(0, r1.getColumn());
		Assert.assertEquals(book.getMaxRowIndex(), r1.getLastRow());
		Assert.assertEquals(book.getMaxColumnIndex(), r1.getLastColumn());
		
		
		r1 = SRanges.range(sheet1,3,4);
		Assert.assertEquals(3, r1.getRow());
		Assert.assertEquals(4, r1.getColumn());
		Assert.assertEquals(3, r1.getLastRow());
		Assert.assertEquals(4, r1.getLastColumn());
		
		r1 = SRanges.range(sheet1,3,4,5,6);
		Assert.assertEquals(3, r1.getRow());
		Assert.assertEquals(4, r1.getColumn());
		Assert.assertEquals(5, r1.getLastRow());
		Assert.assertEquals(6, r1.getLastColumn());
	}
	
	
	@Test
	public void testGeneralCellValue1(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = book.createSheet("Sheet 1");
		Date now = new Date();
		ErrorValue err = new ErrorValue(ErrorValue.INVALID_FORMULA);
		SCell cell = sheet.getCell(1, 1);
		
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		
		SRanges.range(sheet,1,1).setEditText("abc");
		Assert.assertEquals(CellType.STRING, cell.getType());
		Assert.assertEquals("abc",cell.getValue());
		
		SRanges.range(sheet,1,1).setEditText("123");
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(123,cell.getNumberValue().intValue());
		
		SRanges.range(sheet,1,1).setEditText("2013/01/01");
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals("2013/01/01",new SimpleDateFormat("yyyy/MM/dd").format((Date)cell.getDateValue()));
		
		SRanges.range(sheet,1,1).setEditText("tRue");
		Assert.assertEquals(CellType.BOOLEAN, cell.getType());
		Assert.assertEquals(Boolean.TRUE,cell.getBooleanValue());
		
		SRanges.range(sheet,1,1).setEditText("FalSe");
		Assert.assertEquals(CellType.BOOLEAN, cell.getType());
		Assert.assertEquals(Boolean.FALSE,cell.getBooleanValue());
		
		SRanges.range(sheet,1,1).setEditText("=SUM(999)");
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals("SUM(999)", cell.getFormulaValue());
		Assert.assertEquals(999D, cell.getValue());
		
		try{
			SRanges.range(sheet,1,1).setEditText("=SUM)((999)");
			Assert.fail("not here");
		}catch(InvalidModelOpException x){
			//old value
			Assert.assertEquals(CellType.FORMULA, cell.getType());
			Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
			Assert.assertEquals("SUM(999)", cell.getFormulaValue());
			Assert.assertEquals(999D, cell.getValue());
		}

		SRanges.range(sheet,1,1).setEditText("");
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertEquals(null,cell.getValue());		
		
	}
	
	@Test
	public void testGeneralCellValue2(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = book.createSheet("Sheet 1");
		Date now = new Date();
		ErrorValue err = new ErrorValue(ErrorValue.INVALID_FORMULA);
		SCell cell = sheet.getCell(1, 1);
		
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		
		SRanges.range(sheet,1,1).setValue("abc");
		Assert.assertEquals(CellType.STRING, cell.getType());
		Assert.assertEquals("abc",cell.getValue());
		
		SRanges.range(sheet,1,1).setValue(123D);
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(123D,cell.getValue());
		
		
		SRanges.range(sheet,1,1).setValue(now);
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(now,cell.getDateValue());
		
		SRanges.range(sheet,1,1).setValue(Boolean.TRUE);
		Assert.assertEquals(CellType.BOOLEAN, cell.getType());
		Assert.assertEquals(Boolean.TRUE,cell.getValue());
		
		
		SRanges.range(sheet,1,1).setValue("=SUM(999)");
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals("SUM(999)", cell.getFormulaValue());
		Assert.assertEquals(999D, cell.getValue());
		
		try{
			SRanges.range(sheet,1,1).setValue("=SUM)((999)");
			Assert.fail("not here");
		}catch(InvalidModelOpException x){
			Assert.assertEquals(CellType.FORMULA, cell.getType());
			Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
			Assert.assertEquals("SUM(999)", cell.getFormulaValue());
			Assert.assertEquals(999D, cell.getValue());
		}
		
		SRanges.range(sheet,1,1).setValue("");
		Assert.assertEquals(CellType.STRING, cell.getType());
		Assert.assertEquals("",cell.getValue());			
	}
	
	@Test
	public void testFormulaDependency(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = book.createSheet("Sheet 1");
		
		SRanges.range(sheet,0,0).setEditText("999");
		SRanges.range(sheet,0,1).setValue("=SUM(A1)");
		
		
		SCell cell = sheet.getCell(0, 0);
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(999, cell.getNumberValue().intValue());
		
		cell = sheet.getCell(0, 1);
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals("SUM(A1)", cell.getFormulaValue());
		Assert.assertEquals(999D, cell.getValue());
		
		final AtomicInteger a0counter = new AtomicInteger(0);
		final AtomicInteger b0counter = new AtomicInteger(0);
		final AtomicInteger unknowcounter = new AtomicInteger(0);
		
		book.addEventListener(new ModelEventListener() {
			public void onEvent(ModelEvent event) {
				if(event.getName().equals(ModelEvents.ON_CELL_CONTENT_CHANGE)){
					CellRegion region = event.getRegion();
					if(region.getRow()==0&&region.getColumn()==0){
						a0counter.incrementAndGet();
					}else if(region.getRow()==0&&region.getColumn()==1){
						b0counter.incrementAndGet();
					}else{
						unknowcounter.incrementAndGet();
					}
				}
			}
		});
		
		SRanges.range(sheet,0,0).setEditText("888");
		Assert.assertEquals(1, b0counter.intValue());
		Assert.assertEquals(1, a0counter.intValue());
		Assert.assertEquals(0, unknowcounter.intValue());
		
		SRanges.range(sheet,0,0).setEditText("777");
		Assert.assertEquals(2, b0counter.intValue());
		Assert.assertEquals(2, a0counter.intValue());
		Assert.assertEquals(0, unknowcounter.intValue());
		
		SRanges.range(sheet,0,0).setEditText("777");//in last update, set edit text is always notify cell change
		Assert.assertEquals(3, b0counter.intValue());
		Assert.assertEquals(3, a0counter.intValue());
		Assert.assertEquals(0, unknowcounter.intValue());
		
		
	}
/* from D7
A	1	4	5	=SUM(E7:F7)
B	2	5	7	=SUM(E8:F8)
C	3	6	9	=SUM(E9:F9)
	
 */
	@Test
	public void testAutoFilterRange(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = book.createSheet("Sheet 1");
		
		SRange range = SRanges.range(sheet,"D4").findAutoFilterRange();
		
		Assert.assertNull(range);
		
		
		sheet.getCell("D7").setValue("A");
		sheet.getCell("D8").setValue("B");
		sheet.getCell("D9").setValue("C");
		sheet.getCell("E7").setValue(1);
		sheet.getCell("E8").setValue(2);
		sheet.getCell("E9").setValue(3);
		
		sheet.getCell("F7").setValue(4);
		sheet.getCell("F8").setValue(5);
		sheet.getCell("F9").setValue(6);
		
		sheet.getCell("G7").setValue("=SUM(E7:F7)");
		sheet.getCell("G8").setValue("=SUM(E8:F8)");
		sheet.getCell("G9").setValue("=SUM(E9:F9)");
		
		Assert.assertEquals(5D, sheet.getCell("G7").getValue());
		Assert.assertEquals(7D, sheet.getCell("G8").getValue());
		Assert.assertEquals(9D, sheet.getCell("G9").getValue());
		
		
		String nullLoc[] = new String[]{"C5","D5","G5","H5","B6","B7","B9","B10","J6","J7","J9","J10","C11","D11","G11","H11"};
		for(String loc:nullLoc){
			Assert.assertNull("at "+loc,SRanges.range(sheet,loc).findAutoFilterRange());
		}
		
		String inSideLoc[] = new String[]{"D7","D8","D9","E7","E8","E9","F7","F8","F9","G7","G8","G9"};
		for(String loc:inSideLoc){
			range = SRanges.range(sheet,loc).findAutoFilterRange();
			String at = "at "+loc;
			Assert.assertNotNull(at,range);
			Assert.assertEquals(at,6, range.getRow());
			Assert.assertEquals(at,3, range.getColumn());
			Assert.assertEquals(at,8, range.getLastRow());
			Assert.assertEquals(at,6, range.getLastColumn());
		}
		
		Object inCornerLoc[][] = new Object[][]{
			new Object[]{"C6",-1,-1,0,0},//area, row,column,lastRow,lastColumn offset
			new Object[]{"C7",0,-1,0,0},
			new Object[]{"C9",0,-1,0,0},
			new Object[]{"C10",0,-1,1,0},
			
			new Object[]{"H6",-1,0,0,1},
			new Object[]{"H7",0,0,0,1},
			new Object[]{"H9",0,0,0,1},
			new Object[]{"H10",0,0,1,1},
			
			new Object[]{"D6",-1,0,0,0},
			new Object[]{"E6",-1,0,0,0},
			new Object[]{"F6",-1,0,0,0},
			new Object[]{"G6",-1,0,0,0},
			
			new Object[]{"D10",0,0,1,0},
			new Object[]{"E10",0,0,1,0},
			new Object[]{"F10",0,0,1,0},
			new Object[]{"G10",0,0,1,0}
		};
		
		for(Object loc[]:inCornerLoc){
			range = SRanges.range(sheet,(String)loc[0]).findAutoFilterRange();
			String at = "at "+loc[0];
			Assert.assertNotNull(at,range);
			Assert.assertEquals(at,6+(Integer)loc[1], range.getRow());
			Assert.assertEquals(at,3+(Integer)loc[2], range.getColumn());
			Assert.assertEquals(at,8+(Integer)loc[3], range.getLastRow());
			Assert.assertEquals(at,6+(Integer)loc[4], range.getLastColumn());
		}

	}
	
	@Test
	public void addPicture(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = book.createSheet("Picture");
		
		assertEquals(0, sheet.getPictures().size());
		
		try {
			AImage zklogo = new AImage(RangeTest.class.getResource("zklogo.png"));
			
			ViewAnchor anchor = new ViewAnchor(0, 1, zklogo.getWidth()/2, zklogo.getHeight()/2);
			SPicture picture = SRanges.range(sheet).addPicture(anchor, zklogo.getByteData(), SPicture.Format.PNG);
			
			assertEquals(1, sheet.getPictures().size());
			assertEquals(Format.PNG, picture.getFormat());
			assertEquals(zklogo.getWidth()/2, picture.getAnchor().getWidth());
			
//			ImExpTestUtil.write(book, Type.XLSX); //human checking
		} catch (IOException e) {
			e.printStackTrace();
			fail();
		}
	}
	
	@Test
	public void deletePicture(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = book.createSheet("Picture");
		
		assertEquals(0, sheet.getPictures().size());
		
		try {
			AImage zklogo = new AImage(RangeTest.class.getResource("zklogo.png"));
			
			ViewAnchor anchor = new ViewAnchor(0, 1, zklogo.getWidth()/2, zklogo.getHeight()/2);
			SPicture picture = SRanges.range(sheet).addPicture(anchor, zklogo.getByteData(), SPicture.Format.PNG);
			
			assertEquals(1, sheet.getPictures().size());
			assertEquals(Format.PNG, picture.getFormat());
			assertEquals(zklogo.getWidth()/2, picture.getAnchor().getWidth());
			
			SRanges.range(sheet).deletePicture(picture);
			assertEquals(0, sheet.getPictures().size());
			
//			ImExpTestUtil.write(book, Type.XLSX); //human checking
		} catch (IOException e) {
			e.printStackTrace();
			fail();
		}
	}
	
	@Test
	public void movePicture(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = book.createSheet("Picture");
		
		assertEquals(0, sheet.getPictures().size());
		
		try {
			AImage zklogo = new AImage(RangeTest.class.getResource("zklogo.png"));
			
			ViewAnchor anchor = new ViewAnchor(0, 1, zklogo.getWidth()/2, zklogo.getHeight()/2);
			SPicture picture = SRanges.range(sheet).addPicture(anchor, zklogo.getByteData(), SPicture.Format.PNG);
			
			assertEquals(1, sheet.getPictures().size());
			assertEquals(Format.PNG, picture.getFormat());
			assertEquals(zklogo.getWidth()/2, picture.getAnchor().getWidth());
			
			ViewAnchor newAnchor = new ViewAnchor(3, 4, zklogo.getWidth()/2, zklogo.getHeight()/2);
			SRanges.range(sheet).movePicture(picture, newAnchor);
			assertEquals(3, picture.getAnchor().getRowIndex());
			assertEquals(4, picture.getAnchor().getColumnIndex());
			
//			ImExpTestUtil.write(book, Type.XLSX); //human checking
		} catch (IOException e) {
			e.printStackTrace();
			fail();
		}
	}
	
	@Test
	public void testDeleteSheet(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");
		
		sheet1.getCell("A1").setValue(11);
		sheet1.getCell("B1").setValue("=A1");
		sheet2.getCell("B1").setValue("=Sheet1!A1");
		Assert.assertEquals(11D, sheet1.getCell("B1").getValue());
		Assert.assertEquals(11D, sheet2.getCell("B1").getValue());
		
		SRanges.range(sheet1).deleteSheet();
		
		Assert.assertEquals("#REF!", sheet2.getCell("B1").getErrorValue().getErrorString());
	}
	
	@Test
	public void testSetHyperlink(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = book.createSheet("Sheet1");
		SRange range = SRanges.range(sheet1,"A1:B2");
		range.setHyperlink(HyperlinkType.URL, "http://www.google.com", "Google");
		
		SHyperlink link = range.getHyperlink();
		Assert.assertEquals(HyperlinkType.URL, link.getType());
		Assert.assertEquals("http://www.google.com", link.getAddress());
		Assert.assertEquals("Google", link.getLabel());
		
		Assert.assertEquals("Google", sheet1.getCell("A1").getStringValue());
		
		link = SRanges.range(sheet1,"B2").getHyperlink();
		Assert.assertEquals(HyperlinkType.URL, link.getType());
		Assert.assertEquals("http://www.google.com", link.getAddress());
		Assert.assertEquals("Google", link.getLabel());
		Assert.assertEquals("Google", sheet1.getCell("B2").getStringValue());
	}
	
	@Test
	public void testSetStyle(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = book.createSheet("Sheet1");
		SCellStyle style1 = book.createCellStyle(true);
		
		SRanges.range(sheet1,"A1:B2").setCellStyle(style1);
		Assert.assertEquals(style1, SRanges.range(sheet1,"A1").getCellStyle());
		Assert.assertEquals(style1, SRanges.range(sheet1,"A2").getCellStyle());
		Assert.assertEquals(style1, SRanges.range(sheet1,"B1").getCellStyle());
		Assert.assertEquals(style1, SRanges.range(sheet1,"B2").getCellStyle());
		Assert.assertEquals(book.getDefaultCellStyle(), SRanges.range(sheet1,"C1").getCellStyle());
		Assert.assertEquals(book.getDefaultCellStyle(), SRanges.range(sheet1,"C2").getCellStyle());
		
	}
	
	@Test
	public void testShiftChartFormula() {
		SBook book = SBooks.createBook("Book1");
		SSheet sheet = book.createSheet("Sheet1");
		SChart chart = sheet.addChart(ChartType.LINE, new ViewAnchor(0, 0, 100, 100));
		SGeneralChartData data = (SGeneralChartData)chart.getData();
		
		String nameF = "Sheet1:A1";
		String areaF = "Sheet1:A2:A3";
		data.setCategoriesFormula(areaF);
		SSeries series = data.addSeries();
		series.setXYZFormula(nameF, areaF, areaF, areaF);

		// extend
		SRanges.range(sheet, "3").getRows().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE); // whole row
		areaF = "Sheet1:A2:A4";
		assertEquals(areaF, data.getCategoriesFormula());
		assertEquals(nameF, series.getNameFormula());
		assertEquals(areaF, series.getXValuesFormula());
		assertEquals(areaF, series.getYValuesFormula());
		assertEquals(areaF, series.getZValuesFormula());
		
		// shrink
		SRanges.range(sheet, "A3:A4").delete(DeleteShift.UP); // area
		areaF = "Sheet1:A2:A2";
		assertEquals(areaF, data.getCategoriesFormula());
		assertEquals(nameF, series.getNameFormula());
		assertEquals(areaF, series.getXValuesFormula());
		assertEquals(areaF, series.getYValuesFormula());
		assertEquals(areaF, series.getZValuesFormula());
		
		// move series name
		SRanges.range(sheet, "1").getRows().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		nameF = "Sheet1:A2";
		areaF = "Sheet1:A3:A3";
		assertEquals(areaF, data.getCategoriesFormula());
		assertEquals(nameF, series.getNameFormula());
		assertEquals(areaF, series.getXValuesFormula());
		assertEquals(areaF, series.getYValuesFormula());
		assertEquals(areaF, series.getZValuesFormula());
 
		// change data to column direction
		nameF = "Sheet1:A1";
		areaF = "Sheet1:B1:F1";
		data.setCategoriesFormula(areaF);
		series.setXYZFormula(nameF, areaF, areaF, areaF);
		
		// extend
		SRanges.range(sheet, "C:D").getColumns().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE); // whole column
		areaF = "Sheet1:B1:H1";
		assertEquals(areaF, data.getCategoriesFormula());
		assertEquals(nameF, series.getNameFormula());
		assertEquals(areaF, series.getXValuesFormula());
		assertEquals(areaF, series.getYValuesFormula());
		assertEquals(areaF, series.getZValuesFormula());
		
		// shrink
		SRanges.range(sheet, "B1:F1").delete(DeleteShift.LEFT);
		areaF = "Sheet1:B1:C1";
		assertEquals(areaF, data.getCategoriesFormula());
		assertEquals(nameF, series.getNameFormula());
		assertEquals(areaF, series.getXValuesFormula());
		assertEquals(areaF, series.getYValuesFormula());
		assertEquals(areaF, series.getZValuesFormula());

		// move series name in column direction
		SRanges.range(sheet, "A1:C1").insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_NONE);
		nameF = "Sheet1:D1";
		areaF = "Sheet1:E1:F1";
		assertEquals(areaF, data.getCategoriesFormula());
		assertEquals(nameF, series.getNameFormula());
		assertEquals(areaF, series.getXValuesFormula());
		assertEquals(areaF, series.getYValuesFormula());
		assertEquals(areaF, series.getZValuesFormula());

		
		// special cases
		
		// delete all
		SRanges.range(sheet, "B1:F1").delete(DeleteShift.LEFT);
		nameF = "Sheet1:#REF!";
		areaF = "Sheet1:#REF!";
		assertEquals(areaF, data.getCategoriesFormula());
		assertEquals(nameF, series.getNameFormula());
		assertEquals(areaF, series.getXValuesFormula());
		assertEquals(areaF, series.getYValuesFormula());
		assertEquals(areaF, series.getZValuesFormula());
	}
}
