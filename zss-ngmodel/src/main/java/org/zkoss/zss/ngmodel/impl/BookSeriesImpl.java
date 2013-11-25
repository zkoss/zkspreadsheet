package org.zkoss.zss.ngmodel.impl;

import java.util.HashMap;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyEngine;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;

public class BookSeriesImpl extends AbstractBookSeries {
	private static final long serialVersionUID = 1L;
	
	final private HashMap<String,AbstractBook> books;
	
	final private DependencyEngine dependencyEngine;
	
	public BookSeriesImpl(AbstractBook book){
		books = new HashMap<String, AbstractBook>(1);
		books.put(book.getBookName(), book);
		dependencyEngine = EngineFactory.getInstance().createDependencyEngine();
	}
	
	public NBook getBook(String name) {
		return books.get(name);
	}

	@Override
	DependencyTable getDependencyTable() {
		return dependencyEngine.getDependencyTable();
	}


}
