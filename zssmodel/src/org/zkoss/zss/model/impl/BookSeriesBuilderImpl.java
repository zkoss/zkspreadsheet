package org.zkoss.zss.model.impl;

import java.util.Set;

import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeriesBuilder;

public class BookSeriesBuilderImpl extends SBookSeriesBuilder {

	@Override
	public void buildBookSeries(Set<SBook> books) {
		buildBookSeries(books.toArray(new SBook[books.size()]));
	}

	@Override
	public void buildBookSeries(SBook... books) {
		//check type
		AbstractBookAdv bookadvs[] = new AbstractBookAdv[books.length];
		int i = 0;
		for(SBook b: books){
			if(!(b instanceof AbstractBookAdv)){
				throw new IllegalStateException("can't support to build a book "+b+" to book series");
			}
			bookadvs[i] = (AbstractBookAdv)b;
			i++;
		}
		new BookSeriesImpl(bookadvs);
	}

}
