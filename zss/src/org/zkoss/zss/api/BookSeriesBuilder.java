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

import java.util.Set;

import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;
import org.zkoss.lang.Strings;
import org.zkoss.zss.api.model.Book;

/**
 * The book series builder which accepts multiple {@link Book} objects makes each of them can reference cells from other books.
 * @author dennis
 * @since 3.0.0
 */
public abstract class BookSeriesBuilder {
	
	private static BookSeriesBuilder _instance;
	
	public static BookSeriesBuilder getInstance(){
		
		if(_instance==null){
			synchronized(BookSeriesBuilder.class){
				if(_instance==null){
					String clz = Library.getProperty("org.zkoss.zss.api.BookSeriesBuilder.class");
					if (!Strings.isEmpty(clz)) {
						try {
							_instance = (BookSeriesBuilder) Classes.forNameByThread(clz).newInstance();
						} catch (Exception e) {
							throw new RuntimeException(e.getMessage(),e);
						}
					}else{
						_instance = new BookSeriesBuilder() {
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
		return _instance;
	}
	
	
	abstract public void buildBookSeries(Set<Book> books);
	abstract public void buildBookSeries(Book... books);
}
