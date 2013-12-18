package org.zkoss.zss.ngmodel;

import static org.junit.Assert.assertEquals;

import java.io.*;
import java.net.URL;
import java.util.Locale;

import org.junit.*;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.impl.imexp.ExcelImportFactory;

/**
 * @author Hawk
 */
public class ImporterTest extends ImExpTestBase {
	
	static private URL fileUrlUnderTest;
	private NImporter importer; 
	
	/**
	 * For exporter test to specify its exported file to test.
	 * @param fileUrl
	 */
	static public void setFileUnderTest(URL fileUrl){
		fileUrlUnderTest = fileUrl;
	}
	
	@BeforeClass
	static public void initialize(){
		try{
			fileUrlUnderTest = ImporterTest.class.getResource("book/import.xlsx");
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
			book = importer.imports(fileUrlUnderTest.openStream(), "XSSFBook");
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
			book = importer.imports(fileUrlUnderTest, "XSSFBook");
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
			book = importer.imports(new File(fileUrlUnderTest.toURI()), "XSSFBook");
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		assertEquals("XSSFBook", book.getBookName());
		assertEquals(7, book.getNumOfSheet());
	}
	
	//content
	@Test
	public void sheetTest() {
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		
		sheetTest(book);
	}	
	
	@Test
	public void sheetProtectionTest() {
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		sheetProtectionTest(book);
	}
	
	@Test
	public void sheetNamedRangeTest() {
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		sheetNamedRangeTest(book);
	}

	@Test
	public void cellValueTest() {
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		
		cellValueTest(book);
	}
	
	@Test
	public void cellStyleTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		cellStyleTest(book);
	}

	@Test
	public void cellBorderTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		cellBorderTest(book);
	}

	@Test
	public void cellFontNameTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		cellFontNameTest(book);
	}
	
	@Test
	public void cellFontStyleTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		cellFontStyleTest(book);
	}
	
	@Test
	public void cellFontColorTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		cellFontColorTest(book);
		
	}
	
	@Test
	public void rowTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		rowTest(book);
	}

	/**
	 * Information technology — Document description and processing languages — 
	 * Office Open XML File Formats — Part 1: Fundamentals and Markup LanguageReference  
	 * 18.8.30 numFmt (Number Format) 
	 */
	@Test
	public void cellFormatTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		cellFormatTest(book);
	}

	/**
	 * Under different locales, TW and US, should import the same format pattern
	 */
	@Test
	public void formatNotDependLocaleTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		NSheet sheet = book.getSheetByName("Format");
		assertEquals("zh_TW", Locales.getCurrent().toString());
		assertEquals("m/d/yyyy", sheet.getCell(1, 4).getCellStyle().getDataFormat());
		
		Locales.setThreadLocal(Locale.US);
		book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		sheet = book.getSheetByName("Format");
		assertEquals("en_US", Locales.getCurrent().toString());
		assertEquals("m/d/yyyy", sheet.getCell(1, 4).getCellStyle().getDataFormat());
	}
	
	@Test
	public void columnTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		columnTest(book);		
	}
	
	/**
	 * import last column that only has column width change but has all empty cells 
	 */
	@Test
	public void lastChangedColumnTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		lastChangedColumnTest(book);
	}
	
	@Test
	public void viewInfoTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		viewInfoTest(book);
	}

	@Test
	public void mergedTest(){
		NBook book = ImExpTestUtil.loadBook(fileUrlUnderTest, "XSSFBook");
		mergedTest(book);
	}
}


