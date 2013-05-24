package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.model.Picture;
import org.zkoss.zss.model.sys.XBook;

public class PictureImpl implements Picture{
	
	ModelRef<XBook> bookRef;
	ModelRef<org.zkoss.poi.ss.usermodel.Picture> picRef;
	
	public PictureImpl(ModelRef<XBook> bookRef, SimpleRef<org.zkoss.poi.ss.usermodel.Picture> picRef) {
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
		PictureImpl other = (PictureImpl) obj;
		if (picRef == null) {
			if (other.picRef != null)
				return false;
		} else if (!picRef.equals(other.picRef))
			return false;
		return true;
	}
	
	public org.zkoss.poi.ss.usermodel.Picture getNative() {
		return picRef.get();
	}
	
	
	public String getId(){
		return getNative().getPictureId();
	}
}
