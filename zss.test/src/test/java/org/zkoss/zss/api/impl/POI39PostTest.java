package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.SheetCtrl;

/**
 * Test case for poi 3.9 migration
 * @author dennis
 *
 */
public class POI39PostTest {

	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		ZssContext.setThreadLocal(new ZssContext(Locale.TAIWAN,-1));
	}
	
	@After
	public void tearDown() throws Exception {
		ZssContext.setThreadLocal(null);
	}
	
	@Test
	public void testReadWriteChartPicutreNumber() throws Exception{
		//only picture/chart
		//chart first
		testReadWriteChartPicutreNumber(Util.loadBook("book/poi39_shape1.xls"),0,1,2);
		//picture fisrt
		testReadWriteChartPicutreNumber(Util.loadBook("book/poi39_shape1.xls"),1,2,1);
		
//		//has another shape at the begin
		testReadWriteChartPicutreNumber(Util.loadBook("book/poi39_shape2.xls"),0,1,2);
//		//has another shape at the middle		
		testReadWriteChartPicutreNumber(Util.loadBook("book/poi39_shape2.xls"),1,2,1);
//		//mixed
		testReadWriteChartPicutreNumber(Util.loadBook("book/poi39_shape2.xls"),2,2,2);
		
	}
	
	private void testReadWriteChartPicutreNumber(Book book,int sheetIndex,int chartNumber,int pictureNmuber) throws Exception{
		Sheet sheet = book.getSheetAt(sheetIndex);
		//in HSSF chart impl, sheet get charts always returns 1.
		//assertEquals(chartNumber,sheet.getCharts().size());
		assertEquals(chartNumber,((SheetCtrl)sheet.getPoiSheet()).getDrawingManager().getChartXs().size());
		assertEquals(pictureNmuber,sheet.getPictures().size());
		
		File f = Setup.getTempFile();
		Exporter exp = Exporters.getExporter();
		exp.export(book, f);
			
		Importer imp = Importers.getImporter();
			
		book = imp.imports(f,book.getBookName());
		sheet = book.getSheetAt(sheetIndex);
		assertEquals(chartNumber,((SheetCtrl)sheet.getPoiSheet()).getDrawingManager().getChartXs().size());
		assertEquals(pictureNmuber,sheet.getPictures().size());
	}
	
	@Test
	public void testBookSheetReference() throws Exception{
		Book book1 = Util.loadBook("book/poi39_ref_book1.xls");
		Book book2 = Util.loadBook("book/poi39_ref_book2.xls");
		BookSeriesBuilder.getInstance().buildBookSeries(book1,book2);
		assertEquals("6",Ranges.range(book1.getSheetAt(0),"B2").getCellFormatText());
	}
	
	@Test
	public void testStyleLimitation2003() throws Exception{
		Book book1 = Util.loadBook("book/blank.xls");
		testStyleLimitation(book1,200,40,3968);//after large than 3968, some style start to be cleaned up
	}
	
//	@Ignore("performance issue")
	@Test
	public void testStyleLimitation2007() throws Exception{
		Book book1 = Util.loadBook("book/blank.xlsx");
		//don't know the 2007 limitation yet, however it has performanec issue when set large among of styles.
//		testStyleLimitation(book1,200,40,3968);
		testStyleLimitation(book1,200,40,200);//don't know the 2007 limitation yet
	}
	
	public void testStyleLimitation(Book book,int rows,int columns,int maxCells) throws Exception{
		
		Sheet sheet1 = book.getSheetAt(0);
		int totalSet = 0;
		
		loop:
		for(int r=0;r<rows;r++){
			for (int c=0;c<columns;c++){
				CellOperationUtil.applyBackgroundColor(Ranges.range(sheet1,r,c),toTestColor(r,c));
//				System.out.println(">>>set "+r+","+c+" "+toTestColor(r,c));
				totalSet++;
				if(totalSet>=maxCells)
					break loop;
			}
		}
		
//		System.out.println(">>>>totalSet "+totalSet);
		
		totalSet = 0;
		loop2:
		for(int r=0;r<rows;r++){
			for (int c=0;c<columns;c++){
				String color = Ranges.range(sheet1,r,c).getCellStyle().getBackgroundColor().getHtmlColor();
				assertEquals("color not equals at row:"+r+",column:"+c+",rxc:"+(r*c),toTestColor(r,c), color);
				totalSet++;
				if(totalSet>=maxCells)
					break loop2;
			}
		}
		
		//export and read again
		File f = Setup.getTempFile();
		Exporter exp = Exporters.getExporter();
		exp.export(book, f);
			
		Importer imp = Importers.getImporter();
		book = imp.imports(f,book.getBookName());
		sheet1 = book.getSheetAt(0);
		
		
		totalSet = 0;
		loop3:
		for(int r=0;r<rows;r++){
			for (int c=0;c<columns;c++){
				String color = Ranges.range(sheet1,r,c).getCellStyle().getBackgroundColor().getHtmlColor();
				assertEquals("color not equals at row:"+r+",column:"+c+",rxc:"+(r*c),toTestColor(r,c), color);
				totalSet++;
				if(totalSet>=maxCells)
					break loop3;
			}
		}
		
	}

	private String toTestColor(int r, int c) {
		StringBuilder sb = new StringBuilder();
		int color = (r<<8)+(c*5);
		sb.insert(0,toColorHex(color));
		sb.insert(0,toColorHex(color>>4));
		sb.insert(0,toColorHex(color>>8));
		sb.insert(0,"#");
		return sb.toString();
	}
	
	public static String toColorHex(int num) {
		num = num & 0xff;
		final String hex = Integer.toHexString(num);
		return hex.length() == 1 ? "0"+hex : hex; 
	}
			
}
