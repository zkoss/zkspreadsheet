package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XSheet;

public class SheetImpl implements Sheet{
	ModelRef<XSheet> sheetRef;
	Book nbook;
	public SheetImpl(ModelRef<XSheet> sheet){
		this.sheetRef = sheet;
	}
	
	public XSheet getNative(){
		return sheetRef.get();
	}
	public ModelRef<XSheet> getRef(){
		return sheetRef;
	}
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((sheetRef == null) ? 0 : sheetRef.hashCode());
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
		SheetImpl other = (SheetImpl) obj;
		if (sheetRef == null) {
			if (other.sheetRef != null)
				return false;
		} else if (!sheetRef.equals(other.sheetRef))
			return false;
		return true;
	}

	public Book getBook() {
		if(nbook!=null){
			return nbook;
		}
		nbook = new BookImpl(new SimpleRef<XBook>(getNative().getBook()));
		return nbook;
	}
	

	public boolean isProtected() {
		return getNative().getProtect();
	}

	public boolean isAutoFilterEnabled() {
		return getNative().isAutoFilterMode();
	}

	public boolean isDisplayGridlines() {
		return getNative().isDisplayGridlines();
	}

	public String getSheetName() {
		return getNative().getSheetName();
	}
	
}
