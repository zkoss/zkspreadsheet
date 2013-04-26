package org.zkoss.zss.essential.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.zkoss.lang.SystemException;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zss.api.NExporter;
import org.zkoss.zss.api.NExporters;
import org.zkoss.zss.api.NImporter;
import org.zkoss.zss.api.NImporters;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.NBook.BookType;


public class BookUtil {

	static public NBook newBook(String bookName, BookType type) {
		try {
			return newBook0(bookName, type);
		} catch (IOException e) {
			throw new SystemException(e.getMessage(), e);
		}
	}

	static private NBook newBook0(String bookName, BookType type)
			throws IOException {
		
		NImporter importer = NImporters.getImporter("excel");
		
		InputStream is = null;
		switch (type) {
		case EXCEL_2003:
			is = BookUtil.class.getResourceAsStream("_blank.xls");
			break;
		case EXCEL_2007:
//			is = BookUtil.class.getResourceAsStream("_blank.xlsx");
//			is = WebApps.getCurrent().getResourceAsStream("/WEB-INF/excelsrc/bug.xlsx");
			is = WebApps.getCurrent().getResourceAsStream("/WEB-INF/excelsrc/copyPasteBase.xlsx");
			break;
		default :
			throw new SystemException("Unknow book type" + type);
		}
		return importer.imports(is,bookName);
	}

	static File workingFolder;

	static public File getWorkingFolder() {
		if (workingFolder == null) {
			synchronized (BookUtil.class) {
				if (workingFolder == null) {
					workingFolder = new File(
							System.getProperty("java.io.tmpdir"), "zsswrk");
					if (!workingFolder.exists()) {
						if (!workingFolder.mkdirs()) {
							throw new SystemException(
									"Can't get working folder:"
											+ workingFolder.getAbsolutePath());
						}
					}
				}
			}
		}
		return workingFolder;
	}

	
	static public String suggestName(NBook book) {
		String bn = book.getBookName();
		BookType type = book.getType();
		
		String ext = type==BookType.EXCEL_2003?".xls":".xlsx";
		int i = bn.lastIndexOf('.');
		if(i==0){
			bn = "book";
		}else if(i>0){
			bn = bn.substring(0,i);
		}
		return bn+ext;
	}

	static public File saveBook(NBook book) throws IOException{
		NExporter exporter = NExporters.getExporter("excel");
		String bn = book.getBookName();
		String ext = book.getType()==BookType.EXCEL_2003?".xls":".xlsx";
		int i = bn.lastIndexOf('.');
		if(i==0){
			bn = "book";
		}else if(i>0){
			bn = bn.substring(0,i);
		}
		
		File f = File.createTempFile(Long.toString(System.currentTimeMillis()),ext,getWorkingFolder());
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(f);
			exporter.export(book, fos);
		}finally{
			if(fos!=null){
				fos.close();
			}
		}
		return f;
	}
}
