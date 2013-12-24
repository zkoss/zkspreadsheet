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
		AbstractBookAdv bookadvs[] = new AbstractBookAdv[books.length];
		int i = 0;
		for(NBook b: books){
			if(!(b instanceof AbstractBookAdv)){
				throw new IllegalArgumentException("can't support to build a book "+b+" to book series");
			}
			bookadvs[i] = (AbstractBookAdv)b;
			i++;
		}
		new BookSeriesImpl(bookadvs);
	}

}
