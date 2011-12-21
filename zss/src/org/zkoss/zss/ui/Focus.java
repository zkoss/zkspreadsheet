/* Created by kindalu 2009/3/2
 * */


package org.zkoss.zss.ui;

import java.lang.ref.WeakReference;

public class Focus {
	/*package*/ final String id;
	/*package*/ String name;
	/*package*/ String color;
	/*package*/ int row,col;
	final private WeakReference<Spreadsheet> ss;
	
	public Focus(String id, String name, String color, int row, int col, Spreadsheet ss) {
		this.id=id;
		this.name=name;
		this.color=color;
		this.row=row;
		this.col=col;
		this.ss= new WeakReference<Spreadsheet>(ss);
	}
	public String getId() {
		return this.id;
	}
	public boolean isDetached() {
		return ss == null || ss.get() == null || ss.get().getDesktop() == null || !ss.get().getDesktop().isAlive();
	}
	public int hashCode() {
		return id == null ? 0 : id.hashCode();
	}
	
	public boolean equals(Object other) {
		if (other == this) {
			return true;
		}
		if (!(other instanceof Focus)) {
			return false;
		}
		
		return  id.equals(((Focus)other).id);
	}
}
