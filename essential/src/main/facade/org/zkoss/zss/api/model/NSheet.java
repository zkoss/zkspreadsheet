package org.zkoss.zss.api.model;

import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Worksheet;

public class NSheet {
	Worksheet sheet;
	public NSheet(Worksheet sheet){
		this.sheet = sheet;
	}
	
	public Object getSyncObject(){
		return sheet;//Rage impl use it to do sync
	}
	
	public Worksheet getNative(){
		return sheet;
	}
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((sheet == null) ? 0 : sheet.hashCode());
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
		if (sheet == null) {
			if (other.sheet != null)
				return false;
		} else if (!sheet.equals(other.sheet))
			return false;
		return true;
	}
	
	
	
	
}
