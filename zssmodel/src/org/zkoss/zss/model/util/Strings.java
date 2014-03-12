package org.zkoss.zss.model.util;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class Strings {

	
	public static boolean isEmpty(String str){
		return str==null||str.length()==0;
	}
	public static boolean isBlank(String str){
		return str==null || str.trim().length()==0;
	}
}
