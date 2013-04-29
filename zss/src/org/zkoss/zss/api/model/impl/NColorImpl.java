package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.zss.api.model.NColor;
import org.zkoss.zss.model.sys.Book;
import org.zkoss.zss.model.sys.impl.BookHelper;

public class NColorImpl implements NColor{

	ModelRef<Book> bookRef;
	ModelRef<Color> colorRef;

	public NColorImpl(ModelRef<Book> book, ModelRef<Color> color) {
		this.bookRef = book;
		this.colorRef = color;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((colorRef == null) ? 0 : colorRef.hashCode());
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
		if (colorRef == null) {
			if (other.colorRef != null)
				return false;
		} else if (!colorRef.equals(other.colorRef))
			return false;
		return true;
	}

	public Color getNative() {
		return colorRef.get();
	}
	public ModelRef<Color> getRef(){
		return colorRef;
	}

	public String getHtmlColor() {
		return BookHelper.colorToHTML(bookRef.get(),colorRef.get());
	}

}
