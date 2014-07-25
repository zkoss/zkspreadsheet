package org.zkoss.zss.model;

import static org.junit.Assert.fail;

import java.io.*;
import java.util.*;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.range.*;
import org.zkoss.zss.range.impl.imexp.*;
import org.zkoss.zss.range.impl.imexp.ExcelExportFactory.Type;

@Ignore("stucked in 621-shared-formula.xlsx")
public class ExcelImporterTest {

	static ExcelImportFactory importFactory = new ExcelImportFactory();
	static 	ExcelExportFactory xlsxExportFactory = new ExcelExportFactory(Type.XLSX);
	static ExcelExportFactory xlsExportFactory = new ExcelExportFactory(Type.XLS);
	
	/**
	 * Import massive Excel files to ensure the importer and exporter can accept various Excel files without crash.
	 * procedure:
	 * 1. import files and export to validate importing.
	 * 2. import previously exported files to validate exporting. 
	 */
	@Test
	public void importExport(){
		File SOURCE_PATH = new File("./src/test/java");
		File EXPORTED_PATH = new File("./target/exported");
		EXPORTED_PATH.mkdir();
		
		List<File> files2Import = new LinkedList<File>();
		LinkedList<File> folders2Scan = new LinkedList<File>();
		
		folders2Scan.add(SOURCE_PATH);
		//find xls and xlsx file in source path
		while(folders2Scan.size()>0){
			File scanningFolder = folders2Scan.removeFirst();
			File[] path2Check = scanningFolder.listFiles();
			for (File  path : path2Check){
				if (path.isDirectory()){
					folders2Scan.addLast(path);
				}else{
					if (path.getName().toLowerCase().endsWith("xls") || path.getName().toLowerCase().endsWith("xlsx")){
						files2Import.add(path);
					}
				}
			}
		}
		

		String workingFile = null;
		Type[] formats = {Type.XLSX, Type.XLS};
		for (Type format : formats){
			String action = null;
			for (File  file : files2Import){
				try{
					//import 
					action = "Fisrt import";
					workingFile = file.getCanonicalPath();
					SBook book = importFactory.createImporter().imports(file, file.getName());

					//export to a specific folder
					String fileName = getFileName(format, book);
					action = "Export "+format;
					getExporter(format).export(book, new File(EXPORTED_PATH.getPath()+File.separator+fileName));
					
					action = "Second import";
					importFactory.createImporter().imports(file, file.getName());
				}catch (Exception e) {
//					System.out.println("failed on ["+action+"] "+workingFile);
					e.printStackTrace();
					fail("failed on ["+action+"] "+workingFile);
				}
			}
			
		}
	}
	
	private SExporter getExporter(Type format){
		if (format == Type.XLSX){
			return xlsxExportFactory.createExporter();
		}else{
			return xlsExportFactory.createExporter();
		}
	}
	
	private String getFileName(Type format, SBook book){
		if (format == Type.XLSX){
			return book.getBookName().endsWith("xls")?book.getBookName()+"x":book.getBookName();
		}else{
			return book.getBookName().endsWith("xls")?book.getBookName():book.getBookName().substring(0, book.getBookName().length()-1);
		}
	}
	
	@SuppressWarnings("unused")
	@Test
	public void debug() throws IOException{
//		File file = new File(".\\src\\test\\java\\Book10.xls");
		File file = new File("./src/test/java/org/zkoss/zss/api/impl/book/280-reapply.xlsx");
		SBook book = new ExcelImportFactory().createImporter().imports(file, "debug");
		File EXPORTED_PATH = new File("./target/exported");
		//getExporter(Type.XLS).export(book, new File(EXPORTED_PATH.getPath()+File.separator+getFileName(Type.XLS, book)));
	}
}
