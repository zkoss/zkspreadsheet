package org.zkoss.zss.api.model;


public interface Sheet {

	public Book getBook();

	public boolean isProtected();

	public boolean isAutoFilterEnabled();

	public boolean isDisplayGridlines();
	
	public boolean isRowHidden(int row);
	
	public boolean isColumnHidden(int column);

	public String getSheetName();

}
