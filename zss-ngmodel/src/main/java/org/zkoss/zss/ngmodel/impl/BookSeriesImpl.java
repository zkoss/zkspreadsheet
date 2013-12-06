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

import java.util.HashMap;
import java.util.concurrent.locks.ReadWriteLock;
import java.util.concurrent.locks.ReentrantReadWriteLock;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyEngine;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class BookSeriesImpl extends BookSeriesAdv {
	private static final long serialVersionUID = 1L;
	
	final private HashMap<String,BookAdv> books;
	
	final private DependencyEngine dependencyEngine;
	
	final private ReadWriteLock lock = new ReentrantReadWriteLock();
	
	public BookSeriesImpl(BookAdv book){
		books = new HashMap<String, BookAdv>(1);
		books.put(book.getBookName(), book);
		dependencyEngine = EngineFactory.getInstance().createDependencyEngine();
	}
	@Override
	public NBook getBook(String name) {
		return books.get(name);
	}

	@Override
	public DependencyTable getDependencyTable() {
		return dependencyEngine.getDependencyTable();
	}
	@Override
	public ReadWriteLock getLock() {
		return lock;
	}
}
