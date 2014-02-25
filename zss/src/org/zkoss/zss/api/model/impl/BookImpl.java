/* BookImpl.java

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
package org.zkoss.zss.api.model.impl;

import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SSheet;
/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class BookImpl implements Book{

	private ModelRef<SBook> _bookRef;
	private BookType _type;
	
	public BookImpl(ModelRef<SBook> ref){
		this._bookRef = ref;
		SBook book = ref.get();
		/*TODO zss 3.5*/
		_type = BookType.XLSX;
//		if (book instanceof HSSFBookImpl) {
//			_type = BookType.XLS;
//		} else if (book instanceof XSSFBookImpl) {
//			_type = BookType.XLSX;
//		} else {
//			throw new IllegalArgumentException("unknow book type "+book);
//		}
	}

	public SBook getNative() {
		return _bookRef.get();
	}
	
	@Override
	public SBook getInternalBook(){
		return _bookRef.get();
	}
	
	public ModelRef<SBook> getRef(){
		return _bookRef;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((_bookRef == null) ? 0 : _bookRef.hashCode());
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		BookImpl other = (BookImpl) obj;
		if (_bookRef == null) {
			if (other._bookRef != null)
				return false;
		} else if (!_bookRef.equals(other._bookRef))
			return false;
		return true;
	}

	public String getBookName() {
		return getNative().getBookName();
	}
	
	public BookType getType(){
		return _type; 
	}

	public int getSheetIndex(Sheet sheet) {
		if(sheet==null) return -1;
		return getNative().getSheetIndex(((SheetImpl)sheet).getNative());
	}

	public int getNumberOfSheets() {
		return getNative().getNumOfSheet();
	}
	
	public SheetImpl getSheetAt(int index){
		SSheet sheet = getNative().getSheet(index);
		return new SheetImpl(_bookRef,new SimpleRef<SSheet>(sheet));
	}
	
	public SheetImpl getSheet(String name){
		SSheet sheet = getNative().getSheetByName(name);
		
		return sheet==null?null:new SheetImpl(_bookRef,new SimpleRef<SSheet>(sheet));
	}

	
	/*TODO zss 3.5
	@Override
	@Deprecated
	public Workbook getPoiBook() {
		return null;
	}
	*/

	@Override
	public void setShareScope(String scope) {
		getNative().setShareScope(scope);
	}

	@Override
	public String getShareScope() {
		return getNative().getShareScope();
	}

	@Override
	@Deprecated
	public Object getSync() {
		return getNative();
	}

	@Override
	public boolean hasNameRange(String name) {
		return getNative().getNameByName(name)!=null;
	}

	@Override
	public int getMaxRows() {
		return getNative().getMaxRowSize();
	}

	@Override
	public int getMaxColumns() {
		return getNative().getMaxColumnSize();
	}

	@Override
	public ReadWriteLock getLock() {
		return getNative().getBookSeries().getLock();
	}
	
}
