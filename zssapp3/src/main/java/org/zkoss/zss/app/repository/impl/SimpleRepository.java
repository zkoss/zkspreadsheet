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
		FileOutputStream fos = null;
		try{
			File f = ((SimpleBookInfo)info).getFile();
			fos = new FileOutputStream(f);
			Exporters.getExporter().export(book, fos);
		}finally{
			if(fos!=null)
				fos.close();
		}
		return info;
	}
	
	public synchronized BookInfo saveAs(String bookname,Book book) throws IOException {
		
		String name = FileUtil.getName(bookname);
		String ext = "";
		switch(book.getType()){
		case EXCEL_2003:
			ext = ".xls";
			break;
		case EXCEL_2007:
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
}
