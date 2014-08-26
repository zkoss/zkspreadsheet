package org.zkoss.zss;

import java.awt.Desktop;
import java.awt.Desktop.Action;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.model.sys.format.FormatEngine;
import org.zkoss.zss.model.sys.format.FormatResult;

/**
 * a helper for testing
 * 
 * @author kuro, Hawk
 * 
 */
public class Util {
	
	public static Book swap(Book book){
		try{
			File t = Setup.getTempFile();
			Exporters.getExporter().export(book, t);
			book = Importers.getImporter().imports(t, book.getBookName());
			return book;
		}catch(Exception x){
			throw new RuntimeException(x.getMessage(),x);
		}
	}

	public static Book loadBook(File file) {
		try {
			return Importers.getImporter().imports(file, file.getName());
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage(),e);
		}
	}
	
	public static Book loadBook(Object base,String respath) {
		if(base==null){
			base = Util.class;
		}
		if(!(base instanceof Class)){
			base = base.getClass();
		}
		
		@SuppressWarnings("rawtypes")
		final InputStream is = ((Class)base).getResourceAsStream(respath);
		try {
			int index = respath.lastIndexOf("/");
			String bookName = index==-1?respath:respath.substring(index+1);
			return Importers.getImporter().imports(is, bookName);
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage(),e);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
				}
			}
		}
	}
	
	public static void export(Book workbook, File file) {
		Exporter excelExporter = Exporters.getExporter();
		try {
			excelExporter.export(workbook, file);
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage(),e);
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
	
	public static DateFormat getDateFormat(String pattern){
		return getDateFormat(pattern,null);
	}
	public static DateFormat getDateFormat(String pattern, Locale locale){
		return locale==null?new SimpleDateFormat(pattern):new SimpleDateFormat(pattern,locale);
	}

	public static Date getDateOnly(Date d) {
		DateFormat df = getDateFormat("yyyy/MM/dd");
		try {
			return df.parse(df.format(d));
		} catch (ParseException e) {
			throw new RuntimeException(e.getMessage(),e);
		}
	}
	
	public static String getFormatHTMLColor(Sheet sheet, int row, int col) {
		SSheet xSheet = sheet.getInternalSheet();
		SCell cell = xSheet.getCell(row,col);
		FormatEngine engine = EngineFactory.getInstance().createFormatEngine();
		FormatResult r = engine.format(cell, new FormatContext(ZssContext.getCurrent().getLocale()));
		return  r.getColor()!=null?r.getColor().getHtmlColor():null;
	}
}
