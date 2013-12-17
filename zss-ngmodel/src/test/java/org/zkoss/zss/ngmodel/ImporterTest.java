package org.zkoss.zss.ngmodel;

import static org.junit.Assert.*;

import java.io.*;
import java.text.DateFormat;
import java.util.Locale;

import org.junit.*;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.impl.imexp.ExcelImportFactory;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NCellStyle.FillPattern;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NFont.TypeOffset;

/**
 * @author Hawk
 */
public class ImporterTest {
	
	static private File fileUnderTest;
	private NImporter importer; 
	
	/**
	 * For exporter test to specify its exported file to test.
	 * @param file
	 */
	static public void setFileUnderTest(File file){
		fileUnderTest = file;
	}
	
	@BeforeClass
	static public void initialize(){
		try{
			fileUnderTest = new File(ImporterTest.class.getResource("book/import.xlsx").toURI());
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	@Before
	public void beforeTest(){
		importer= new ExcelImportFactory().createImporter();
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	//API
	
	@Test
	public void importByInputStream(){
		InputStream streamUnderTest = null;
		NBook book = null;
		try {
			streamUnderTest = new FileInputStream(fileUnderTest);
			book = importer.imports(streamUnderTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}finally{
			if(streamUnderTest!=null){
				try{
					streamUnderTest.close();
				}catch(Exception e){
					e.printStackTrace();
				}
			}
		}
		
		assertEquals("XSSFBook", book.getBookName());
		assertEquals(7, book.getNumOfSheet());
	}
	
	@Test
	public void importByUrl(){
		NBook book = null;
		try {
			book = importer.imports(fileUnderTest.toURL(), "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		assertEquals("XSSFBook", book.getBookName());
		assertEquals(7, book.getNumOfSheet());
	}
	
	@Test
	public void importByFile() {
		NBook book = null;
		try {
			book = importer.imports(fileUnderTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		assertEquals("XSSFBook", book.getBookName());
		assertEquals(7, book.getNumOfSheet());
	}
	
	//content
	@Test
	public void sheetTest() {
		NBook book = importBook(fileUnderTest, "XSSFBook");
		
		assertEquals(7, book.getNumOfSheet());

		NSheet sheet1 = book.getSheetByName("Value");
		assertEquals("Value", sheet1.getSheetName());
		assertEquals(21, sheet1.getDefaultRowHeight());
		assertEquals(72, sheet1.getDefaultColumnWidth());
		
		NSheet sheet2 = book.getSheetByName("Style");
		assertEquals("Style", sheet2.getSheetName());
		NSheet sheet3 = book.getSheetByName("NamedRange");
		assertEquals("NamedRange", sheet3.getSheetName());
	}	
	
	@Test
	public void sheetProtectionTest() {
		NBook book = importBook(fileUnderTest, "XSSFBook");
		assertFalse(book.getSheetByName("Value").isProtected());
		assertTrue(book.getSheetByName("sheet-protection").isProtected());
	}
	
	@Test
	public void sheetNamedRangeTest() {
		NBook book = importBook(fileUnderTest, "XSSFBook");
		assertEquals(2, book.getNumOfName());
		assertEquals("NamedRange!$B$2:$D$3", book.getNameByName("TestRange1", "NamedRange").getRefersToFormula());
		assertEquals("NamedRange!$F$2", book.getNameByName("RangeMerged", "NamedRange").getRefersToFormula());
	}

	@Test
	public void cellValueTest() {
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("Value");
		//text
		assertEquals(NCell.CellType.STRING, sheet.getCell(0,1).getType());
		assertEquals("B1", sheet.getCell(0,1).getStringValue());
		assertEquals("C1", sheet.getCell(0,2).getStringValue());
		assertEquals("D1", sheet.getCell(0,3).getStringValue());
		
		//number
		assertEquals(NCell.CellType.NUMBER, sheet.getCell(1,1).getType());
		assertEquals(123, sheet.getCell(1,1).getNumberValue().intValue());
		assertEquals(123.45, sheet.getCell(1,2).getNumberValue().doubleValue(), 0.01);
		
		//date
		assertEquals(NCell.CellType.NUMBER, sheet.getCell(2,1).getType());
		assertEquals(41618, sheet.getCell(2,1).getNumberValue().intValue());
		assertEquals("Dec 10, 2013", DateFormat.getDateInstance(DateFormat.MEDIUM, Locale.US).format(sheet.getCell(2,1).getDateValue()));
		assertEquals(0.61, sheet.getCell(2,2).getNumberValue().doubleValue(), 0.01);
		assertEquals("2:44:10 PM", DateFormat.getTimeInstance (DateFormat.MEDIUM, Locale.US).format(sheet.getCell(2,2).getDateValue()));
		
		//formula
		assertEquals(NCell.CellType.FORMULA, sheet.getCell(3,1).getType());
		assertEquals("SUM(10,20)", sheet.getCell(3,1).getFormulaValue());
		assertEquals("ISBLANK(B1)", sheet.getCell(3,2).getFormulaValue());
		assertEquals("B1", sheet.getCell(3,3).getFormulaValue());
		
		//error
		assertEquals(NCell.CellType.ERROR, sheet.getCell(4,1).getType());
		assertEquals(ErrorValue.INVALID_NAME, sheet.getCell(4,1).getErrorValue().getCode());
		assertEquals(ErrorValue.INVALID_VALUE, sheet.getCell(4,2).getErrorValue().getCode());
		
		//blank
		assertEquals(NCell.CellType.BLANK, sheet.getCell(5,1).getType());
		assertEquals("", sheet.getCell(5,1).getStringValue());
	}
	
	@Test
	public void cellStyleTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("Style");
		assertEquals(true, sheet.getCell(24, 1).getCellStyle().isWrapText());
		//alignment
		assertEquals(VerticalAlignment.TOP, sheet.getCell(26, 1).getCellStyle().getVerticalAlignment());
		assertEquals(VerticalAlignment.CENTER, sheet.getCell(26, 2).getCellStyle().getVerticalAlignment());
		assertEquals(VerticalAlignment.BOTTOM, sheet.getCell(26, 3).getCellStyle().getVerticalAlignment());
		assertEquals(Alignment.LEFT, sheet.getCell(27, 1).getCellStyle().getAlignment());
		assertEquals(Alignment.CENTER, sheet.getCell(27, 2).getCellStyle().getAlignment());
		assertEquals(Alignment.RIGHT, sheet.getCell(27, 3).getCellStyle().getAlignment());
		//cell filled color
		assertEquals("#FF0000", sheet.getCell(11, 0).getCellStyle().getFillColor().getHtmlColor());
		assertEquals("#00FF00", sheet.getCell(11, 1).getCellStyle().getFillColor().getHtmlColor());
		assertEquals("#0000FF", sheet.getCell(11, 2).getCellStyle().getFillColor().getHtmlColor());

		//ensure cell style reusing
		assertTrue(sheet.getCell(27, 0).getCellStyle().equals(sheet.getCell(26, 0).getCellStyle()));
		assertTrue(sheet.getCell(28, 0).getCellStyle().equals(sheet.getCell(26, 0).getCellStyle()));
		
		//fill pattern
		assertEquals(FillPattern.SOLID_FOREGROUND, sheet.getCell(37, 1).getCellStyle().getFillPattern());
		assertEquals(FillPattern.ALT_BARS, sheet.getCell(37, 2).getCellStyle().getFillPattern());
		
		NSheet protectedSheet = book.getSheetByName("sheet-protection");
		assertEquals(true, protectedSheet.getCell(0, 0).getCellStyle().isLocked());
		assertEquals(false, protectedSheet.getCell(1, 0).getCellStyle().isLocked());
		
		
	}

	@Test
	public void cellBorderTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("cell-border");
		assertEquals(BorderType.NONE, sheet.getCell(2, 0).getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, sheet.getCell(2, 1).getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, sheet.getCell(2, 2).getCellStyle().getBorderTop());
		assertEquals(BorderType.THIN, sheet.getCell(2, 3).getCellStyle().getBorderLeft());
		assertEquals(BorderType.THIN, sheet.getCell(2, 4).getCellStyle().getBorderRight());
		
		assertEquals(BorderType.THIN, sheet.getCell(4, 1).getCellStyle().getBorderBottom());
		assertEquals(BorderType.DOTTED, sheet.getCell(4, 2).getCellStyle().getBorderBottom());
		assertEquals(BorderType.DASHED, sheet.getCell(4, 3).getCellStyle().getBorderBottom());
		
		assertEquals("#FF0000", sheet.getCell(14, 1).getCellStyle().getBorderBottomColor().getHtmlColor());
		assertEquals("#0000FF", sheet.getCell(14, 1).getCellStyle().getBorderLeftColor().getHtmlColor());
		assertEquals("#0000FF", sheet.getCell(14, 1).getCellStyle().getBorderTopColor().getHtmlColor());
		assertEquals("#FF0000", sheet.getCell(14, 1).getCellStyle().getBorderRightColor().getHtmlColor());
	}

	@Test
	public void cellFontNameTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("Style");
		assertEquals("Arial", sheet.getCell(3, 0).getCellStyle().getFont().getName());
		assertEquals("Arial Black", sheet.getCell(3, 1).getCellStyle().getFont().getName());
		assertEquals("Calibri", sheet.getCell(3, 2).getCellStyle().getFont().getName());
	}
	
	@Test
	public void cellFontStyleTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("Style");
		assertEquals(NFont.Boldweight.BOLD, sheet.getCell(9, 0).getCellStyle().getFont().getBoldweight());
		assertTrue(sheet.getCell(9, 1).getCellStyle().getFont().isItalic());
		assertTrue(sheet.getCell(9, 2).getCellStyle().getFont().isStrikeout());
		assertEquals(NFont.Underline.SINGLE, sheet.getCell(9, 3).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.DOUBLE, sheet.getCell(9, 4).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.SINGLE_ACCOUNTING, sheet.getCell(9, 5).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.DOUBLE_ACCOUNTING, sheet.getCell(9, 6).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.NONE, sheet.getCell(9, 7).getCellStyle().getFont().getUnderline());
		
		//height
		assertEquals(8, sheet.getCell(6, 0).getCellStyle().getFont().getHeightPoints());
		assertEquals(72, sheet.getCell(6, 3).getCellStyle().getFont().getHeightPoints());
		
		//type offset
		assertEquals(TypeOffset.SUPER, sheet.getCell(32, 1).getCellStyle().getFont().getTypeOffset());
		assertEquals(TypeOffset.SUB, sheet.getCell(32, 2).getCellStyle().getFont().getTypeOffset());
		assertEquals(TypeOffset.NONE, sheet.getCell(32, 3).getCellStyle().getFont().getTypeOffset());
	}
	
	@Test
	public void cellFontColorTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("Style");
		assertEquals("#000000", sheet.getCell(0, 0).getCellStyle().getFont().getColor().getHtmlColor());
		assertEquals("#FF0000", sheet.getCell(1, 0).getCellStyle().getFont().getColor().getHtmlColor());
		assertEquals("#00FF00", sheet.getCell(1, 1).getCellStyle().getFont().getColor().getHtmlColor());
		assertEquals("#0000FF", sheet.getCell(1, 2).getCellStyle().getFont().getColor().getHtmlColor());
		
	}
	
	@Test
	public void rowTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("Style");
		assertEquals(28, sheet.getRow(0).getHeight());
		assertEquals(20, sheet.getRow(1).getHeight());
		//style
		NCellStyle rowStyle1 = sheet.getRow(34).getCellStyle();
		assertEquals("#0000FF",rowStyle1.getFont().getColor().getHtmlColor());
		assertEquals(12,rowStyle1.getFont().getHeightPoints());
		NCellStyle rowStyle2 = sheet.getRow(35).getCellStyle();
		assertEquals(true,rowStyle2.getFont().isItalic());
		assertEquals(14,rowStyle2.getFont().getHeightPoints());		
		
		NSheet rowSheet = book.getSheetByName("column-row");
		assertTrue(rowSheet.getRow(9).isHidden());
		assertTrue(rowSheet.getRow(10).isHidden());
	}

	@Test
	public void cellFormatTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("Format");
		assertEquals("#,##0.00", sheet.getCell(1, 1).getCellStyle().getDataFormat());
		assertEquals("\"NT$\"#,##0.00", sheet.getCell(1, 2).getCellStyle().getDataFormat());
		assertEquals("yyyy/m/d", sheet.getCell(1, 4).getCellStyle().getDataFormat());
		//actual "h:mm AM/PM"
//		assertEquals("hh:mm AM/PM", sheet.getCell(1, 5).getCellStyle().getDataFormat());
		assertEquals("0.0%", sheet.getCell(1, 6).getCellStyle().getDataFormat());
		assertEquals("# ??/??", sheet.getCell(3, 1).getCellStyle().getDataFormat());
		assertEquals("0.00E+00", sheet.getCell(3, 2).getCellStyle().getDataFormat());
		assertEquals("@", sheet.getCell(3, 3).getCellStyle().getDataFormat());
		//TODO what characters required to escape
//		assertEquals("[<=9999999]###\\-####;\\(0#\\)\\ ###\\-####", sheet.getCell(3, 4).getCellStyle().getDataFormat());
		
	}
	
	@Test
	public void dataFormatOnLocaleTest(){
//		NBook book = importBook(fileUnderTest, "XSSFBook");
//		NSheet sheet = book.getSheetByName("Format");
//		Locales.setThreadLocal(locale)
//		BuiltinFormats.getBuiltinFormat(index, ZssContext.getCurrent().getLocale())
//		assertEquals("yyyy/m/d", sheet.getCell(1, 4).getCellStyle().getDataFormat());
	}
	
	@Test
	public void columnTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("column-row");
		//column style
		assertEquals(NFont.Boldweight.BOLD, sheet.getColumn(0).getCellStyle().getFont().getBoldweight());
		//width
		assertEquals(228, sheet.getColumn(0).getWidth()); 
		assertEquals(100, sheet.getColumn(1).getWidth());
		assertEquals(102, sheet.getColumn(2).getWidth());
		assertEquals(0, sheet.getColumn(4).getWidth());		//the hidden column
		assertEquals(72, sheet.getColumn(5).getWidth());	//default width
		
		//the hidden column
		assertFalse(sheet.getColumn(3).isHidden());
		assertTrue(sheet.getColumn(4).isHidden());		
	}
	
	/**
	 * import last column that only has column width change but has all empty cells 
	 */
	@Test
	public void lastChangedColumnTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("column-row");
		assertEquals(80, sheet.getColumn(13).getWidth());
	}
	
	@Test
	public void viewInfoTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("column-row");
		assertEquals(3, sheet.getViewInfo().getNumOfRowFreeze());
		assertEquals(1, sheet.getViewInfo().getNumOfColumnFreeze());
		
		//grid line display
		assertTrue(sheet.getViewInfo().isDisplayGridline());
		assertFalse(book.getSheetByName("cell-border").getViewInfo().isDisplayGridline());
	}

	@Test
	public void mergedTest(){
		NBook book = importBook(fileUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("cell-border");
		assertEquals(4, sheet.getMergedRegions().size());
		assertEquals("C31:D33", sheet.getMergedRegions().get(0).getReferenceString());
		assertEquals("B31:B33", sheet.getMergedRegions().get(1).getReferenceString());
		assertEquals("E28:G28", sheet.getMergedRegions().get(2).getReferenceString());
		assertEquals("B28:C28", sheet.getMergedRegions().get(3).getReferenceString());
	}
	
	
	protected NBook importBook(File file, String bookName){
		NBook book = null;
		try {
			book = importer.imports(file, bookName);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return book;
	}
}


