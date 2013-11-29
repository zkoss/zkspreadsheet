package org.zkoss.zss.ngmodel;

public interface NComment {

	public String getText();
	public void setText(String text);
	
	public void setRichText(NRichText text);
	public NRichText setRichText();
	public NRichText getRichText();
	
	public String getAuthor();
	public void setAuthor(String author);
}
