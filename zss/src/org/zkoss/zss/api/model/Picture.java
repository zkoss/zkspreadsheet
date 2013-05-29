/* Picture.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api.model;

/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public interface Picture {
	
	public enum Format{
	    EMF,
	    WMF,
	    PICT,
	    JPEG,
	    PNG,
	    DIB		
	}
	
	public String getId();
}
