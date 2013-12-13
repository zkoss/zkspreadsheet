package org.zkoss.zss.ngmodel.impl;

import java.util.Set;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeriesBuilder;

public class BookSeriesBuilderImpl extends NBookSeriesBuilder {

	@Override
	public void buildBookSeries(Set<NBook> books) {
		buildBookSeries(books.toArray(new NBook[books.size()]));
	}

	@Override
	public void buildBookSeries(NBook... books) {
		//check type
		BookAdv bookadvs[] = new BookAdv[books.length];
		int i = 0;
		for(NBook b: books){
			if(!(b instanceof BookAdv)){
				throw new IllegalArgumentException("can't support to build a book "+b+" to book series");
			}
			bookadvs[i] = (BookAdv)b;
			i++;
		}
		new BookSeriesImpl(bookadvs);
	}

}
