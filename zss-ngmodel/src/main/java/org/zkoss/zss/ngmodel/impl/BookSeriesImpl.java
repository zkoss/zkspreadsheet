/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel.impl;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.concurrent.locks.ReadWriteLock;
import java.util.concurrent.locks.ReentrantReadWriteLock;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.impl.sys.DependencyTableAdv;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class BookSeriesImpl extends BookSeriesAdv {
	private static final long serialVersionUID = 1L;
	
	final private HashMap<String,BookAdv> books;
	
	final private DependencyTable dependencyTable;
	
	final private ReadWriteLock lock = new ReentrantReadWriteLock();
	
	public BookSeriesImpl(BookAdv... books){
		this.books = new LinkedHashMap<String, BookAdv>(1);
		dependencyTable = EngineFactory.getInstance().createDependencyTable();
		for(BookAdv book:books){
			this.books.put(book.getBookName(), book);
			((DependencyTableAdv) dependencyTable)
					.merge((DependencyTableAdv) ((BookSeriesAdv) book
							.getBookSeries()).getDependencyTable());
			book.setBookSeries(this);
		}
	}
	@Override
	public NBook getBook(String name) {
		return books.get(name);
	}

	@Override
	public DependencyTable getDependencyTable() {
		return dependencyTable;
	}
	@Override
	public ReadWriteLock getLock() {
		return lock;
	}
	@Override
	public List<NBook> getBooks() {
		return Collections.unmodifiableList(new ArrayList<NBook>(books.values()));
	}
}
