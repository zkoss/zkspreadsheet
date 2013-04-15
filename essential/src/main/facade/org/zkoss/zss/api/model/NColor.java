package org.zkoss.zss.api.model;

import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;

public class NColor {

	Book book;
	Color color;

	public NColor(Book book, Color color) {
		this.book = book;
		this.color = color;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((color == null) ? 0 : color.hashCode());
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
		if (color == null) {
			if (other.color != null)
				return false;
		} else if (!color.equals(other.color))
			return false;
		return true;
	}

	public Color getNative() {
		return color;
	}

	public String toHtmlColor() {
		return BookHelper.colorToHTML(book,color);
	}

}
