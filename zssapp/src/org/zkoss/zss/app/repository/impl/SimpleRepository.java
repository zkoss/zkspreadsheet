/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.repository.impl;

import java.io.File;
import java.io.FileFilter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.app.repository.BookInfo;
import org.zkoss.zss.app.repository.BookRepository;
import org.zkoss.zss.app.ui.UiUtil;
/**
 * 
 * @author dennis
 *
 */
public class SimpleRepository implements BookRepository{
	File root;
	public SimpleRepository(File root){
		this.root = root;
	}
	
	
	public synchronized List<BookInfo> list() {
		List<BookInfo> books = new ArrayList<BookInfo>();
		for(File f:root.listFiles(new FileFilter() {
			public boolean accept(File file) {
				if(file.isFile() && !file.isHidden()){
					String ext = FileUtil.getNameExtension(file.getName()).toLowerCase();
					if("xls".equals(ext) || "xlsx".equals(ext)){
						return true;
					}
				}
				return false;
			}
		})){
			books.add(new SimpleBookInfo(f,f.getName(),new Date(f.lastModified())));
		}
		return books;
	}

	public synchronized Book load(BookInfo info) throws IOException {
		Book book = Importers.getImporter().imports(((SimpleBookInfo)info).getFile(), info.getName());
		return book;
	}

	public synchronized BookInfo save(BookInfo info, Book book) throws IOException {
		if(UiUtil.isRepositoryReadonly()){
			return null;
		}
		FileOutputStream fos = null;
		try{
			File f = ((SimpleBookInfo)info).getFile();
			//write to temp file first to avoid write error damage original file 
			File temp = File.createTempFile("temp", f.getName());
			fos = new FileOutputStream(temp);
			Exporters.getExporter().export(book, fos);
			
			fos.close();
			fos = null;
			
			FileUtil.copy(temp,f);
			temp.delete();
			
		}finally{
			if(fos!=null)
				fos.close();
		}
		return info;
	}
	
	public synchronized BookInfo saveAs(String bookname,Book book) throws IOException {
		if(UiUtil.isRepositoryReadonly()){
			return null;
		}
		String name = FileUtil.getName(bookname);
		String ext = "";
		switch(book.getType()){
		case XLS:
			ext = ".xls";
			break;
		case XLSX:
			ext = ".xlsx";
			break;
		default:
			throw new RuntimeException("unknow book type");
		}
		File f = new File(root,name+ext);
		int c = 0;
		if(f.exists()){
			f = new File(root,name+"("+(++c)+")"+ext);
		}
		SimpleBookInfo info = new SimpleBookInfo(f,f.getName(),new Date());
		return save(info,book);
	}


	public boolean delete(BookInfo info) throws IOException {
		if(UiUtil.isRepositoryReadonly()){
			return false;
		}
		File f = ((SimpleBookInfo)info).getFile();
		if(!f.exists()){
			return false;
		}
		return f.delete();
	}
}
