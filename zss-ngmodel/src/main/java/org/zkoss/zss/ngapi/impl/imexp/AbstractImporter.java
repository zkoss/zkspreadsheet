/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl.imexp;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.ErrorConstants;
import org.zkoss.poi.xssf.usermodel.XSSFCell;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;

/**
 * 
 * @author Hawk
 * @since 3.5.0 
 */
public abstract class AbstractImporter implements NImporter{

	@Override
	public NBook imports(File file, String bookName) throws IOException {
		InputStream is = null;
		try{
			is = new BufferedInputStream(new FileInputStream(file));
			return imports(is,bookName);
		}finally{
			if(is!=null){
				try{
					is.close();
				}catch(Exception e){
					e.printStackTrace();
				}
			}
		}
	}

	@Override
	public NBook imports(URL url, String bookName) throws IOException {
		InputStream is = null;
		try {
			is = url.openStream();
			return imports(is, bookName);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
	}
}
