package org.zkoss.zss.api.impl;

import java.awt.Desktop;
import java.awt.Desktop.Action;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.text.ParseException;

import junit.framework.Assert;

import org.apache.commons.math.complex.Complex;
import org.zkoss.poi.ss.formula.functions.ComplexFormat;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.model.Book;

/**
 * a helper for testing
 * 
 * @author kuro, Hawk
 * 
 */
public class Util {

	public static boolean isAMergedRange(Range range) {
		org.zkoss.poi.ss.usermodel.Sheet sheet = range.getSheet().getPoiSheet();
		// go through all region
		for (int number = sheet.getNumMergedRegions(); number > 0; number--) {
			org.zkoss.poi.ss.util.CellRangeAddress addr = sheet.getMergedRegion(number - 1);
			// match four corner
			if (addr.getFirstRow() == range.getRow() && addr.getLastRow() == range.getLastRow() && addr.getFirstColumn() == range.getColumn() && addr.getLastColumn() == range.getLastColumn()) {
				return true;
			}
		}
		return false; // doesn't match any region
	}

	public static Book loadBook(String filename) {
		final InputStream is = Util.class.getResourceAsStream(filename);
		try {
			return Importers.getImporter().imports(is, filename);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
				}
			}
		}
		return null;
	}

	public static void export(Book workbook, String filename) {
		Exporter excelExporter = Exporters.getExporter("excel");
		FileOutputStream fos;
		try {
			fos = new FileOutputStream(new File(ShiftTest.class.getResource("").getPath() + filename));
			excelExporter.export(workbook, fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * Let OS open the file for human eye checking.
	 * 
	 * @param file
	 * @throws IOException
	 */
	public static void open(File file) {
		if (Desktop.isDesktopSupported()) {
			if (Desktop.getDesktop().isSupported(Action.OPEN)) {
				try {
					Desktop.getDesktop().open(file);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	public static void open(String filePath) {
		try {
			open(new File(Util.class.getResource(filePath).toURI()));
		} catch (URISyntaxException e) {
			e.printStackTrace();
		}
	}

}
