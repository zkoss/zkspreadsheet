package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.zss.model.Book;

public class NPicture {
	
	
	public enum Format{
	    EMF,
	    WMF,
	    PICT,
	    JPEG,
	    PNG,
	    DIB		
	}
	
	ModelRef<Book> bookRef;
	ModelRef<Picture> picRef;
	
	public NPicture(ModelRef<Book> bookRef, SimpleRef<Picture> picRef) {
		this.bookRef = bookRef;
		this.picRef = picRef;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((picRef == null) ? 0 : picRef.hashCode());
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
		NColor other = (NColor) obj;
		if (picRef == null) {
			if (other.colorRef != null)
				return false;
		} else if (!picRef.equals(other.colorRef))
			return false;
		return true;
	}
	
	public Picture getNative() {
		return picRef.get();
	}
}
