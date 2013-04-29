package org.zkoss.zss.api.model;


public interface NSheet {

	public NBook getBook();

	public boolean isProtected();

	public boolean isAutoFilterEnabled();

	public boolean isDisplayGridlines();

	public String getSheetName();

}
