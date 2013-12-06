/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface NComment {

	public String getText();
	public void setText(String text);
	
	public void setRichText(NRichText text);
	public NRichText setRichText();
	public NRichText getRichText();
	
	public String getAuthor();
	public void setAuthor(String author);
}
