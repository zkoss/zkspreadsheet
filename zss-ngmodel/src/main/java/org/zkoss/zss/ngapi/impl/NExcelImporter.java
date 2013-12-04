package org.zkoss.zss.ngapi.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.impl.BookImpl;

public class NExcelImporter implements NImporter{


	@Override
	public NBook imports(File file, String bookName) throws IOException {
		InputStream is = null;
		try{
			is = new FileInputStream(file);
			return imports(is,bookName);
		}finally{
			if(is!=null){
				try{
					is.close();
				}catch(Exception x){};
			}
		}
	}

	@Override
	public NBook imports(URL url, String bookName) throws IOException {
		InputStream is = null;
		try{
			is = url.openStream();
			return imports(is,bookName);
		}finally{
			if(is!=null){
				try{
					is.close();
				}catch(Exception x){};
			}
		}
	}
	

	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {
		NBook book = new BookImpl(bookName);
		//implement here
		return book;
	}

}
