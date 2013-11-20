package org.zkoss.zss.ngmodel.util;

public class Validations {

	
	public static void argNotNull(Object obj){
		argNotNull("argument is null",obj);
	}
	public static void argNotNull(String message, Object... obj){
		if(obj==null){
			throw new IllegalArgumentException(message);
		}
		for(int i=0;i<obj.length;i++){
			if(obj[i]==null){
				throw new IllegalArgumentException(message +" "+ i);
			}
		}
	}
	
	public static void argInstance(Object obj,Class clz){
		argInstance("can't cast to ",obj,clz);
	}
	public static void argInstance(String message,Object obj,Class clz){
		if(obj!=null && !clz.isAssignableFrom(obj.getClass())){
			throw new IllegalArgumentException(message+" "+clz);
		}
	}
	
}
