/* Color.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 5:24:22 PM , Created by sam
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
public class Color {
	
	String hex;
	
	public Color(String color) {
		color = color.toUpperCase();
		if (isRGBA(color)) {
			this.hex = rgbaTohex(color);
		} else if (isRGB(color)) {
			this.hex = rgbToHex(color);
		} else {
			if (color.indexOf("#") < 0)
				throw new IllegalArgumentException("color shall be rgb format or hex format, not: " + color);
			this.hex = color;
		}
	}
	
	public String getColor() {
		return hex;
	}
	
	private boolean isRGB(String color) {
		return color.indexOf("RGB") >= 0;
	}
	
	private boolean isRGBA(String color) {
		return color.indexOf("RGBA") >= 0;
	}
	
	private boolean isHex(String color) {
		boolean hex = color.indexOf("#") >= 0;
		boolean rgb = color.indexOf("RGB") >= 0;
		if (hex || rgb)
			return hex;
		else
			throw new IllegalArgumentException("[" + color + "] is not color format");
	}
	// convert RGBA color data to hex
	private String rgbaTohex(String rgba) {
		int index = rgba.indexOf("RGBA");
		if (index < 0)
			throw new IllegalArgumentException("[" + rgba + "] is not RGBA format");
		
		rgba = rgba.substring(index + "RGBA".length());
		rgba = rgba.replace("(", "");
		rgba = rgba.replace(")", "");
		String[] ary = rgba.split(",");
		int r = Integer.parseInt(ary[0].trim());
		int g = Integer.parseInt(ary[1].trim());
		int b = Integer.parseInt(ary[2].trim());
		int a = Integer.parseInt(ary[3].trim());
		
	    if (r > 255 || g > 255 || b > 255 || a > 255)
	    	throw new IllegalArgumentException();
	    
	    return "#" + toHex(r)+toHex(g)+toHex(b);
	}

	//assume UpperCase
	private String rgbToHex(String rgb) {
		int index = rgb.indexOf("RGB");
		if (index < 0)
			throw new IllegalArgumentException("[" + rgb + "] is not RGB format");
		
		rgb = rgb.substring(index + "RGB".length());
		rgb = rgb.replace("(", "");
		rgb = rgb.replace(")", "");
		String[] ary = rgb.split(",");
		int r = Integer.parseInt(ary[0].trim());
		int g = Integer.parseInt(ary[1].trim());
		int b = Integer.parseInt(ary[2].trim());
		
		return "#" + toHex(r)+toHex(g)+toHex(b);
	}
	
	protected String toHex(int n) {
		 //Integer num = Integer.parseInt(n, 10);
		 if (n == 0) return "00";
		 n = Math.max(0, Math.min(n, 255));
		 return "" + "0123456789ABCDEF".charAt((n - n % 16)/16)
		      + "0123456789ABCDEF".charAt(n%16);
	}

	@Override
	public boolean equals(Object obj) {
		if (obj instanceof Color) {
			Color that = (Color)obj;
			return Objects.equal(this.hex, that.hex);
		}
		return false;
	}

	@Override
	public String toString() {
		return Objects.toStringHelper(this)
			.add("color", hex)
			.toString();
	}
}