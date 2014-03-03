package org.zkoss.zss.model;

import static org.junit.Assert.fail;

import java.io.*;
import java.net.URL;

import org.zkoss.zss.model.SBook;
import org.zkoss.zss.range.SExporter;
import org.zkoss.zss.range.SImporter;
import org.zkoss.zss.range.impl.imexp.*;
import org.zkoss.zss.range.impl.imexp.ExcelExportFactory.Type;

/**
 * Utility for write and load book
 * @author kuro
 *
 */
public class ImExpTestUtil {
	

	static public String DEFAULT_EXPORT_TARGET_PATH = "./target/";
	static public String DEFAULT_EXPORT_FILE_NAME_XLSX = "exported.xlsx";
	static public String DEFAULT_EXPORT_FILE_NAME_XLS = "exported.xls";
	
	public static File write(SBook book, Type type) {
		if (type.equals(ExcelExportFactory.Type.XLSX)){
			return writeBookToFile(book, new File(DEFAULT_EXPORT_TARGET_PATH + DEFAULT_EXPORT_FILE_NAME_XLSX), type);
		}else{
			return writeBookToFile(book, new File(DEFAULT_EXPORT_TARGET_PATH + DEFAULT_EXPORT_FILE_NAME_XLS), type);
		}
	}
	
	public static File writeBookToFile(SBook book, String filePath, Type type) {
		return writeBookToFile(book, new File(filePath), type);
	}
	
	public static File writeBookToFile(SBook book, File outFile, Type type) {
		try {
			outFile = new File(DEFAULT_EXPORT_TARGET_PATH + outFile.getName());
			outFile.createNewFile();
			SExporter exporter = new ExcelExportFactory(type).createExporter();
			exporter.export(book, outFile);
		} catch (Exception e) {
			e.printStackTrace();
			fail(e.toString());
		}
		return outFile;
	}
	
	/**
	 * Utility to load book from input stream
	 * @param is
	 * @return
	 */
	public static SBook loadBook(InputStream is) {
		SImporter importer = new ExcelImportFactory().createImporter();
		SBook book = null;
		try {
			book = importer.imports(is, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				is.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return book;
	}
	
	/**
	 * Utility to load book from file
	 * @param file
	 * @param bookName TODO
	 * @return
	 */
	public static SBook loadBook(File file, String bookName) {
		try {
			return new ExcelImportFactory().createImporter().imports(file, bookName);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	public static SBook loadBook(URL fileUrl, String bookName) {
		try {
			return new ExcelImportFactory().createImporter().imports(fileUrl, bookName);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}

}
