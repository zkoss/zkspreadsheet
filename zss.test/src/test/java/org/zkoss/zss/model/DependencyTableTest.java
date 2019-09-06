package org.zkoss.zss.model;

import org.junit.*;
import org.junit.rules.Stopwatch;
import org.zkoss.util.Locales;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.model.impl.SimpleBookSeriesImpl;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.range.*;
import org.zkoss.zss.range.impl.imexp.ExcelImportFactory;
import org.zkoss.zss.zats.PrintStopwatch;

import java.io.IOException;
import java.util.*;

import static org.junit.Assert.assertEquals;
import static org.zkoss.zss.model.FormulaEvalTest.toCellRef;


public class DependencyTableTest {

	private SImporter importer;

	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}

	@Before
	public void beforeTest() {
		importer= new ExcelImportFactory().createImporter();
		Locales.setThreadLocal(Locale.TAIWAN);
	}

	@Test
	public void formulaRef() throws IOException {
		String fileName = "book/formula-ref.xlsx";
		SBook book = importer.imports(this.getClass().getResourceAsStream(fileName), fileName);

		Set<Ref> dependents = ((SimpleBookSeriesImpl) book.getBookSeries()).getDependencyTable().getDependents(toCellRef(book.getBookName(), book.getSheet(0).getSheetName(), null, "A1"));
		assertEquals(100, dependents.size());
	}

	@Test
	public void editCells() throws IOException {
		String fileName = "book/formula-ref.xlsx";
		SBook book = importer.imports(this.getClass().getResourceAsStream(fileName), fileName);

		long startTime = System.currentTimeMillis();
		int startingRow = 0;
		int nCells = 25;
		SSheet sheet1 = book.getSheet(0);
		for (int row = startingRow; row < startingRow + nCells ; row++ ) {
			SRanges.range(sheet1, row, 0).setEditText("0");
		}
		System.out.println((System.currentTimeMillis() - startTime) + " ms");
		assertEquals(SRanges.range(sheet1, "K1").getCellFormatText(), "2925");
		assertEquals(SRanges.range(sheet1, "L1").getCellFormatText(), "11.7");
		assertEquals(SRanges.range(sheet1, "M1").getCellFormatText(), "7.89");
		assertEquals(SRanges.range(sheet1, "N1").getCellFormatText(), "8555824.15");
	}

	@Test
	public void editCells2() throws IOException {
		String fileName = "book/formula-9000.xlsx";
		SBook book = importer.imports(this.getClass().getResourceAsStream(fileName), fileName);

		long startTime = System.currentTimeMillis();
		int startingRow = 4;
		int nCells = 10;
		for (int row = startingRow ; row < startingRow + nCells ; row++ ) {
			SRanges.range(book.getSheet(0), row, 7).setEditText("1");
		}
		System.out.println((System.currentTimeMillis() - startTime) + " ms");
	}

	@Rule
	public final Stopwatch stopwatch = new PrintStopwatch();
}
