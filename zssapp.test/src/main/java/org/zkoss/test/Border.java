/* Border.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 5:19:44 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import com.google.common.base.Objects;

/**
 * @author sam
 *
 */
public class Border {
	
	String width;
	String style;
	Color color;
	
	public Border(String width, String style, String color) {
		this.width = width;
		this.style = style.toLowerCase();
		this.color = new Color(color);
	}
	
	public String getWidth() {
		return width;
	}
	
	public String getStyle() {
		return style;
	}
	
	public Color getColor() {
		return color;
	}

	@Override
	public boolean equals(Object obj) {
		if (obj instanceof Border) {
			Border that = (Border)obj;
			return Objects.equal(width, that.width)
				&& Objects.equal(style, that.style)
				&& Objects.equal(color, that.color);
		}
		return false;
	}

	@Override
	public String toString() {
		return Objects.toStringHelper(this)
			.add("width", width)
			.add("style", style)
			.add("color", color)
			.toString();
	}
	
	
}