package org.zkoss.zss.api.impl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.model.Book;

/**
 * a helper for testing
 * @author kuro
 *
 */
public class Util {
	
	public static boolean isAMergedRange(Range range) {
		org.zkoss.poi.ss.usermodel.Sheet sheet = range.getSheet().getPoiSheet();
		// go thorugh all region
		for(int number = sheet.getNumMergedRegions(); number > 0; number--) {
			org.zkoss.poi.ss.util.CellRangeAddress addr = sheet.getMergedRegion(number-1);
			// match four corner
			if(addr.getFirstRow() == range.getRow() && addr.getLastRow() == range.getLastRow()
					&& addr.getFirstColumn() == range.getColumn() && addr.getLastColumn() == range.getLastColumn()) {
				return true;
			}
		}
		return false; // doesn't match any region
	}

	public static Book loadBook(String filename) throws IOException {
		final InputStream is = Util.class.getResourceAsStream(filename);
		return Importers.getImporter().imports(is, filename);
	}
	
	public static void export(Book workbook, String filename) throws IOException {
		Exporter excelExporter = Exporters.getExporter("excel");
		FileOutputStream fos = new FileOutputStream(new File(ShiftTest.class.getResource("").getPath() + filename));
		excelExporter.export(workbook, fos);
	}
}
