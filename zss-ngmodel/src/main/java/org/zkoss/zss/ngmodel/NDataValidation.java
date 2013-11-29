package org.zkoss.zss.ngmodel;

public interface NDataValidation {
	public enum ErrorStyle {
		STOP, WARNING, INFO;
	}
	
	public ErrorStyle getErrorStyle();
	public void setErrorStyle(ErrorStyle errorStyle);
}
