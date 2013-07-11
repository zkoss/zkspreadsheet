/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.repository.impl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

import org.zkoss.lang.Strings;
import org.zkoss.lang.SystemException;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Book.BookType;
import org.zkoss.zss.app.repository.BookInfo;
import org.zkoss.zss.app.repository.BookRepository;
/**
 * 
 * @author dennis
 *
 */
public class BookUtil {

	/**
	 * Gets suggested file name of a book
	 */
	static public String suggestFileName(String name,BookRepository rep){
		int i = 0;
		name = FileUtil.getName(name);
		String ext = FileUtil.getNameExtension(name);
		
		Set<String> names = new HashSet<String>();

		for(BookInfo info:rep.list()){
			names.add(info.getName());
		}
		String sname = name+"."+ext;
		while(names.contains(sname)){
			sname = name+"("+ ++i +")."+ext;
		}
		return sname;
	}
	
	/**
	 * Gets suggested file name of a book
	 */
	static public String suggestFileName(Book book) {
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
	
	static File workingFolder;
	
	static public File getWorkingFolder() {
		if (workingFolder == null) {
			synchronized (BookUtil.class) {
				if (workingFolder == null) {
					workingFolder = new File(
							System.getProperty("java.io.tmpdir"), "zssappwork");
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
	
	static public File saveBookToWorkingFolder(Book book) throws IOException{
		Exporter exporter = Exporters.getExporter();
		String bn = suggestFileName(book);
		
		String name = FileUtil.getName(bn);
		String ext = FileUtil.getNameExtension(bn);
		
		if(Strings.isBlank(name)){
			name = "book";
		}
		
		
		File f = File.createTempFile(name,"."+ext,getWorkingFolder());
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
