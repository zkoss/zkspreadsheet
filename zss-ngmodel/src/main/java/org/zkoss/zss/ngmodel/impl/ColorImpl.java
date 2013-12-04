package org.zkoss.zss.ngmodel.impl;

import java.util.Arrays;

import org.zkoss.zss.ngmodel.NColor;

public class ColorImpl extends ColorAdv {
	private static final long serialVersionUID = 1L;
	private final byte[] rgb;

	public static final ColorAdv WHITE = new ColorImpl("#FFFFFF");
	public static final ColorAdv BLACK = new ColorImpl("#000000");
	public static final ColorAdv RED = new ColorImpl("#FF0000");
	public static final ColorAdv GREEN = new ColorImpl("#00FF00");
	public static final ColorAdv BLUE = new ColorImpl("#0000FF");

	public ColorImpl(byte[] rgb) {
		if (rgb == null) {
			throw new IllegalArgumentException("null rgb array");
		} else if (rgb.length != 3) {
			throw new IllegalArgumentException("wrong rgb length");
		}
		this.rgb = rgb;
	}

	public ColorImpl(byte r, byte g, byte b) {
		this.rgb = new byte[] { r, g, b };
	}

	public ColorImpl(String htmlColor) {
		final int offset = htmlColor.charAt(0) == '#' ? 1 : 0;
		final short red = Short.parseShort(
				htmlColor.substring(offset + 0, offset + 2), 16); // red
		final short green = Short.parseShort(
				htmlColor.substring(offset + 2, offset + 4), 16); // green
		final short blue = Short.parseShort(
				htmlColor.substring(offset + 4, offset + 6), 16); // blue
		final byte r = (byte) (red & 0xff);
		final byte g = (byte) (green & 0xff);
		final byte b = (byte) (blue & 0xff);
		this.rgb = new byte[] { r, g, b };
	}
	
	private static final char HEX[] = {
		'0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F'
	};

	@Override
	public String getHtmlColor() {
		StringBuilder sb = new StringBuilder("#");
		for(byte c:rgb){
			int n = c & 0xff;
			sb.append(HEX[n/16]);//high
			sb.append(HEX[n%16]);//low
		}
		return sb.toString();
	}

	@Override
	public byte[] getRGB() {
		byte[] c = new byte[rgb.length];
		System.arraycopy(rgb, 0, c, 0, rgb.length);
		return c;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + Arrays.hashCode(rgb);
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		ColorImpl other = (ColorImpl) obj;
		if (!Arrays.equals(rgb, other.rgb))
			return false;
		return true;
	}

	
}
