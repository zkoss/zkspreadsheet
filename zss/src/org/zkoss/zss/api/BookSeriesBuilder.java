/* BookSeriesBuilder.java

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
package org.zkoss.zss.api;

import java.util.List;
import java.util.Set;

import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;
import org.zkoss.lang.Strings;
import org.zkoss.zss.api.model.Book;

/**
 * The book series builder
 * @author dennis
 *
 */
public abstract class BookSeriesBuilder {
	
	static BookSeriesBuilder instance;
	
	public static BookSeriesBuilder getInstance(){
		
		if(instance==null){
			synchronized(BookSeriesBuilder.class){
				if(instance==null){
					String clz = Library.getProperty("org.zkoss.zss.api.BookSeriesBuilder.class");
					if (!Strings.isEmpty(clz)) {
						try {
							instance = (BookSeriesBuilder) Classes.forNameByThread(clz).newInstance();
						} catch (Exception e) {
							throw new RuntimeException(e.getMessage(),e);
						}
					}else{
						instance = new BookSeriesBuilder() {
							@Override
							public void buildBookSeries(Set<Book> books) {
								throw new RuntimeException("not implemented");
							}
							@Override
							public void buildBookSeries(Book[] books) {
								throw new RuntimeException("not implemented");
							}
						};
					}
				}
			}
		}
		return instance;
	}
	
	
	abstract public void buildBookSeries(Set<Book> books);
	abstract public void buildBookSeries(Book... books);
}
