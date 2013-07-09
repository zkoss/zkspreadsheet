package org.zkoss.zss.app.repository;

import java.util.HashSet;
import java.util.Set;

public class BookUtil {

	
	public static String suggestBookName(String name){
		BookRepository rep = BookRepositoryFactory.getInstance().getRepository();
		int i = 0;
		String sname = name;
		
		Set<String> names = new HashSet<String>();

		for(BookInfo info:rep.list()){
			names.add(info.getName());
		}
		
		while(names.contains(sname)){
			sname = name+"("+ ++i +")";
		}
		return sname;
	}
}
