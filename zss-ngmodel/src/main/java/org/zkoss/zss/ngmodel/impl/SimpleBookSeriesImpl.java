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
import java.util.List;
import java.util.concurrent.locks.ReadWriteLock;
import java.util.concurrent.locks.ReentrantReadWriteLock;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class SimpleBookSeriesImpl extends BookSeriesAdv {
	private static final long serialVersionUID = 1L;
	
	final private BookAdv book;
	List<NBook> array;
	
	final private DependencyTable dependencyTable;
	
	final private ReadWriteLock lock = new ReentrantReadWriteLock();
	
	public SimpleBookSeriesImpl(BookAdv book){
		this.book = book;
		dependencyTable = EngineFactory.getInstance().createDependencyTable();
	}
	@Override
	public NBook getBook(String name) {
		return book.getBookName().equals(name)?book:null;
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
		if(array!=null){
			return array;
		}
		array = new ArrayList<NBook>(1);
		array.add(book);
		return array = Collections.unmodifiableList(array);
	}
}
