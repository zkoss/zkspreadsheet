package org.zkoss.zss.issue;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.Sheet;

public class Issue820Test {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}

	/**
	 * parse a defined name which contains question mark '?'
	 * @throws IOException 
	 */
	@Test
	public void reorderSheet() throws IOException {
		Book book = Util.loadBook(this, "book/820-reorder-sheet.xlsx");
		Sheet sheet1 = book.getSheetAt(0);
		Sheet sheet2 = book.getSheetAt(1);
		Sheet sheet3 = book.getSheetAt(2);
		Sheet sheet4 = book.getSheetAt(3);
		Sheet sheet5 = book.getSheetAt(4);
		Range rngB2 = Ranges.range(sheet1, "B2");
		CellData b2 = rngB2.getCellData();
		Assert.assertEquals("Formula", "SUM(Sheet2:Sheet5!A1)", b2.getFormulaValue());
		Assert.assertEquals("Value a Number Type", CellData.CellType.NUMERIC, b2.getResultType());
		Assert.assertEquals("Value", 14d, b2.getDoubleValue(), 0d);

		//Sheet2 move to be after Sheet5:
		// [Sheet1, Sheet2, Sheet3, Sheet4, Sheet5] 
		// [Sheet1, Sheet3, Sheet4, Sheet5, Sheet2]
		Range rngSheet2 = Ranges.range(sheet2);
		rngSheet2.setSheetOrder(4);
		
		rngB2 = Ranges.range(sheet1, "B2");
		b2 = rngB2.getCellData();
		Assert.assertEquals("Formula", "SUM(Sheet3:Sheet5!A1)", b2.getFormulaValue());
		Assert.assertEquals("Value a Number Type", CellData.CellType.NUMERIC, b2.getResultType());
		Assert.assertEquals("Value", 12d, b2.getDoubleValue(), 0d);
		
		//Sheet5 move to be before Sheet3:  
		// [Sheet1, Sheet3, Sheet4, Sheet5, Sheet2]
		// [Sheet1, Sheet5, Sheet3, Sheet4, Sheet2]
		Range rngSheet5 = Ranges.range(sheet5);
		rngSheet5.setSheetOrder(1);
		
		rngB2 = Ranges.range(sheet1, "B2");
		b2 = rngB2.getCellData();
		Assert.assertEquals("Formula", "SUM(Sheet3:Sheet4!A1)", b2.getFormulaValue());
		Assert.assertEquals("Value a Number Type", CellData.CellType.NUMERIC, b2.getResultType());
		Assert.assertEquals("Value", 7d, b2.getDoubleValue(), 0d);
		
//		//export and test again
//		File temp = Setup.getTempFile();
//		Exporters.getExporter().export(book, temp);
//		book = Importers.getImporter().imports(temp, "test");
//		Sheet sheet1 = book.getSheetAt(0);
//		String printArea = sheet1.getInternalSheet().getPrintSetup().getPrintArea();
//		Assert.assertEquals("Print Area", "Sheet1!$B$2:$D$6", printArea);
	}
}