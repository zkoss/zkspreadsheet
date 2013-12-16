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
public interface NPicture {

	public enum Format{
		PNG,
		JPG,
		GIF;
		public String getFileExtension() {
			return name().toLowerCase();
		}
		
		
	}
	public NSheet getSheet();
	
	public String getId();
	
	public Format getFormat();
	
	public byte[] getData();
	
	public NViewAnchor getAnchor();
}
