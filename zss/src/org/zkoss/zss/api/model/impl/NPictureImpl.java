package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.zss.api.model.NPicture;
import org.zkoss.zss.model.sys.XBook;

public class NPictureImpl implements NPicture{
	
	ModelRef<XBook> bookRef;
	ModelRef<Picture> picRef;
	
	public NPictureImpl(ModelRef<XBook> bookRef, SimpleRef<Picture> picRef) {
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
		NColorImpl other = (NColorImpl) obj;
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
