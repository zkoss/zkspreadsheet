/* ImporterImpl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.impl.BookImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XImporter;

/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class ImporterImpl implements Importer{
	XImporter importer;
	public ImporterImpl(XImporter importer) {
		this.importer = importer;
	}

	
	public Book imports(InputStream is, String bookName)throws IOException{
		return new BookImpl(new SimpleRef<XBook>(importer.imports(is, bookName)));
	}


	public XImporter getNative() {
		return importer;
	}


	@Override
	public Book imports(File file, String bookName) throws IOException {
		FileInputStream is = null;
		try{
			is = new FileInputStream(file);
			return imports(is,bookName);
		}finally{
			if(is!=null){
				try{
					is.close();
				}catch(Exception x){}//eat
			}
		}
	}


	@Override
	public Book imports(URL url, String bookName) throws IOException {
		InputStream is = null;
		try{
			is = url.openStream();
			return imports(is,bookName);
		}finally{
			if(is!=null){
				try{
					is.close();
				}catch(Exception x){}//eat
			}
		}
	}
}
