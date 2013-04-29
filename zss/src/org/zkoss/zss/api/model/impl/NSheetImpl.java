package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.model.sys.Book;
import org.zkoss.zss.model.sys.Worksheet;

public class NSheetImpl implements NSheet{
	ModelRef<Worksheet> sheetRef;
	NBook nbook;
	public NSheetImpl(ModelRef<Worksheet> sheet){
		this.sheetRef = sheet;
	}
	
	public Worksheet getNative(){
		return sheetRef.get();
	}
	public ModelRef<Worksheet> getRef(){
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
		NSheetImpl other = (NSheetImpl) obj;
		if (sheetRef == null) {
			if (other.sheetRef != null)
				return false;
		} else if (!sheetRef.equals(other.sheetRef))
			return false;
		return true;
	}

	public NBook getBook() {
		if(nbook!=null){
			return nbook;
		}
		nbook = new NBookImpl(new SimpleRef<Book>(getNative().getBook()));
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
