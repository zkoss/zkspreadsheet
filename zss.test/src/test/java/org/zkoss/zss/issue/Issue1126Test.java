package org.zkoss.zss.issue;

import java.util.Locale;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

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
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.api.Exporters;

public class Issue1126Test {
	
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

	@Test
	public void testDateInput(){
		testDateInput1(Util.loadBook(this, "book/blank.xlsx"));
		testDateInput1(Util.loadBook(this, "book/blank.xls"));
	}
	public void testDateInput1(Book book){
		Sheet sheet = book.getSheetAt(0);
		Range r;
		
		//TW
		Setup.setZssLocale(Locale.TAIWAN);
		r = Ranges.range(sheet,"A1");
		Assert.assertEquals("General", r.getCellDataFormat());
		r.setCellEditText("2013-10-25"); // data format in TW
		Assert.assertEquals("yyyy/m/d", r.getCellDataFormat()); //format change
		Assert.assertEquals("2013/10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText());

		r = Ranges.range(sheet,"A2");
		Assert.assertEquals("General", r.getCellDataFormat());
		r.setCellEditText("=VALUE(\"2015-09-05\")"); // data format in TW
		Assert.assertEquals("General", r.getCellDataFormat()); //format change
		Assert.assertEquals("42252", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getResultType());
		Assert.assertEquals("=VALUE(\"2015-09-05\")", r.getCellEditText());

		r = Ranges.range(sheet,"A3");
		Assert.assertEquals("General", r.getCellDataFormat());
		r.setCellEditText("13-10-25"); // data format in TW
		Assert.assertEquals("yyyy/m/d", r.getCellDataFormat()); //format change
		Assert.assertEquals("2013/10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText());

		
		//US
		Setup.setZssLocale(Locale.US);
		
		r = Ranges.range(sheet,"B1");
		Assert.assertEquals("General", r.getCellDataFormat());
		r.setCellEditText("2013-10-25"); // data format in US
		Assert.assertEquals("m/d/yyyy", r.getCellDataFormat()); //format change
		Assert.assertEquals("10/25/2013", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("10/25/2013", r.getCellEditText());

		r = Ranges.range(sheet,"B2");
		Assert.assertEquals("General", r.getCellDataFormat());
		r.setCellEditText("=VALUE(\"2015-09-05\")"); // data format in US
		Assert.assertEquals("General", r.getCellDataFormat()); //format change
		Assert.assertEquals("42252", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getResultType());
		Assert.assertEquals("=VALUE(\"2015-09-05\")", r.getCellEditText());

		r = Ranges.range(sheet,"B3");
		Assert.assertEquals("General", r.getCellDataFormat());
		r.setCellEditText("13-10-25"); // data format in US
		Assert.assertEquals("General", r.getCellDataFormat());//format doesn't change
		Assert.assertEquals("13-10-25", r.getCellFormatText());
		Assert.assertEquals(CellType.STRING, r.getCellData().getType());//it is string
		Assert.assertEquals("13-10-25", r.getCellEditText());
		
		//UK
		Setup.setZssLocale(Locale.UK);
		
		r = Ranges.range(sheet,"C1");
		Assert.assertEquals("General", r.getCellDataFormat());
		r.setCellEditText("2013-10-25"); // data format in TW
		Assert.assertEquals("dd-mm-yyyy", r.getCellDataFormat()); //format change
		Assert.assertEquals("25-10-2013", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("25-10-2013", r.getCellEditText());
		
		r = Ranges.range(sheet,"C2");
		Assert.assertEquals("General", r.getCellDataFormat());
		r.setCellEditText("=VALUE(\"2015-09-05\")"); // data format in TW
		Assert.assertEquals("General", r.getCellDataFormat()); //format change
		Assert.assertEquals("42252", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getResultType());
		Assert.assertEquals("=VALUE(\"2015-09-05\")", r.getCellEditText());
		
		r = Ranges.range(sheet,"C3");
		Assert.assertEquals("General", r.getCellDataFormat());
		r.setCellEditText("13-10-25"); // data format in TW
		Assert.assertEquals("dd-mm-yyyy", r.getCellDataFormat()); //format change
		Assert.assertEquals("13-10-2025", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("13-10-2025", r.getCellEditText());
	}
}
