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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
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
public class SimpleBookSeriesImpl extends AbstractBookSeriesAdv {
	private static final long serialVersionUID = 1L;
	
	final private AbstractBookAdv book;
	List<NBook> array;
	
	final private DependencyTable dependencyTable;
	
	final private ReadWriteLock lock = new ReentrantReadWriteLock();
	
	private transient Map<String, Object> attributes;
	
	public SimpleBookSeriesImpl(AbstractBookAdv book){
		this.book = book;
		dependencyTable = EngineFactory.getInstance().createDependencyTable();
		((DependencyTableAdv)dependencyTable).setBookSeries(this);
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

	@Override
	public Object getAttribute(String name) {
		return name != null ? getAttributeMap().get(name) : null;
	}

	@Override
	public Object setAttribute(String name, Object value) {
		if(name != null) {
			Map<String, Object> map = getAttributeMap();
			return value != null ? map.put(name, value) : map.remove(name);
		} else {
			return null;
		}
	}

	@Override
	public Map<String, Object> getAttributes() {
		return Collections.unmodifiableMap(getAttributeMap());
	}
	
	private Map<String, Object> getAttributeMap() {
		if(attributes == null) {
			attributes = new LinkedHashMap<String, Object>();
		}
		return attributes;
	}
}
