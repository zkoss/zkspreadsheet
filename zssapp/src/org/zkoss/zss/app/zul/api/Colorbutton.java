/* Colorbutton.java

{{IS_NOTE
	Purpose:

	Description:

	History:
		Oct 19, 2010 3:30:26 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.zul.api;

/**
 * @author Sam
 *
 */
public interface Colorbutton {
	/**
	 * Sets the color.
	 * @param color in #RRGGBB format (hexdecimal).
	 */
	public String getColor();
	/**
	 * Returns the color in int
	 * <p>Default: 0x000000
	 */
	public int getRGB();
	/**
	 * Returns the color (in string as #RRGGBB).
	 * <p>Default: #000000
	 */
	public void setColor(String color);

	/** Returns the image URI of the button
	 * <p>Default: null.
	 */
	public String getImage();

	/** Sets the image URI.
	 */
	public void setImage(String src);
}
