/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.repository.impl;
/**
 * 
 * @author dennis
 *
 */
public class FileUtil {

	public static String getNameExtension(String filename){
		int i = filename.lastIndexOf('.');
		if (i > 0) {
		    return filename.substring(i+1);
		}
		return "";
	}
	
	public static String getName(String filename){
		int i = filename.lastIndexOf('.');
		if (i > 0) {
		    return filename.substring(0,i);
		}
		return filename;
	}
}
