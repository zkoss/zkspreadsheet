package org.zkoss.zss.api;

import java.util.List;

import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;
import org.zkoss.lang.Strings;
import org.zkoss.zss.api.model.Book;

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
							public void buildBookSeries(List<Book> books) {
								throw new RuntimeException("not implemented");
							}
						};
					}
				}
			}
		}
		return instance;
	}
	
	
	abstract public void buildBookSeries(List<Book> books);
}
