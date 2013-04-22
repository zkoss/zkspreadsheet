package org.zkoss.zss.api.model;

import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Worksheet;

public class NSheet {
	ModelRef<Worksheet> sheetRef;
	NBook nbook;
	public NSheet(ModelRef<Worksheet> sheet){
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
		NSheet other = (NSheet) obj;
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
		nbook = new NBook(new SimpleRef<Book>(getNative().getBook()));
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
