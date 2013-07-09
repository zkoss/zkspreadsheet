package org.zkoss.zss.app.repository.impl;

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
