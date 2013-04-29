package org.zkoss.zss.api.model;


public interface Sheet {

	public Book getBook();

	public boolean isProtected();

	public boolean isAutoFilterEnabled();

	public boolean isDisplayGridlines();

	public String getSheetName();

}
