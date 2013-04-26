package org.zkoss.zss.api.ui;

import org.zkoss.zss.api.model.ModelRef;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.api.model.SimpleRef;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;

public class NSpreadsheet {
	Spreadsheet ss;

	public NSpreadsheet(Spreadsheet ss) {
		this.ss = ss;
	}
	
	public void setBook(NBook book) {
		ss.setBook(book==null?null:book.getNative());
	}
	public Spreadsheet getNative(){
		return ss;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((ss == null) ? 0 : ss.hashCode());
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
		NSpreadsheet other = (NSpreadsheet) obj;
		if (ss == null) {
			if (other.ss != null)
				return false;
		} else if (!ss.equals(other.ss))
			return false;
		return true;
	}

	public NBook getBook() {
		return new NBook(new SimpleRef<Book>(ss.getBook()));
	}

	public Rect getSelection() {
		return ss.getSelection();
	}

	public NSheet getSelectedSheet() {
		Worksheet sheet = ss.getSelectedSheet();
		return sheet==null?null:new NSheet(new SimpleRef<Worksheet>(sheet));
	}

	public void setHighlight(Rect rect) {
		ss.setHighlight(rect);
	}

	public int getMaxVisibleColumns() {
		return ss.getMaxcolumns();
	}
	public int getMaxVisibleRows() {
		return ss.getMaxrows();
	}
}
