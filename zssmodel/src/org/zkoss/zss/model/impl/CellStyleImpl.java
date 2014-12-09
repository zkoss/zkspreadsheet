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
package org.zkoss.zss.model.impl;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.ByteArrayOutputStream;

import javax.imageio.ImageIO;

import org.apache.commons.codec.binary.Base64;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class CellStyleImpl extends AbstractCellStyleAdv {
	private static final long serialVersionUID = 1L;

	private AbstractFontAdv _font;
	private SColor _fillColor = ColorImpl.WHITE;
	private SColor _backColor = ColorImpl.WHITE;
	private FillPattern _fillPattern = FillPattern.NONE;
	private Alignment _alignment = Alignment.GENERAL;
	private VerticalAlignment _verticalAlignment = VerticalAlignment.BOTTOM;
	private boolean _wrapText = false;

	private BorderType _borderLeft = BorderType.NONE;
	private BorderType _borderTop = BorderType.NONE;
	private BorderType _borderRight = BorderType.NONE;
	private BorderType _borderBottom = BorderType.NONE;
	private SColor _borderTopColor = ColorImpl.BLACK;
	private SColor _borderLeftColor = ColorImpl.BLACK;
	private SColor _borderBottomColor = ColorImpl.BLACK;
	private SColor _borderRightColor = ColorImpl.BLACK;

	private String _dataFormat = FORMAT_GENERAL;
	private boolean _directFormat = false;
	private boolean _locked = true;// default locked as excel.
	private boolean _hidden = false;
	
	private String _patternHtml; //ZSS-841. cached html string for fill pattern

	public CellStyleImpl(AbstractFontAdv font){
		this._font = font;
	}
	
	public SFont getFont(){
		return _font;
	}
	
	public void setFont(SFont font){
		Validations.argInstance(font, AbstractFontAdv.class);
		this._font = (AbstractFontAdv)font;
	}

	@Override
	public SColor getFillColor() {
		return _fillColor;
	}

	@Override
	public void setFillColor(SColor fillColor) {
		Validations.argNotNull(fillColor);
		this._fillColor = fillColor;
		_patternHtml = null; //clear cache
	}

	@Override
	public FillPattern getFillPattern() {
		return _fillPattern;
	}

	@Override
	public void setFillPattern(FillPattern fillPattern) {
		Validations.argNotNull(fillPattern);
		this._fillPattern = fillPattern;
		_patternHtml = null; //clear cache
	}

	@Override
	public Alignment getAlignment() {
		return _alignment;
	}

	@Override
	public void setAlignment(Alignment alignment) {
		Validations.argNotNull(alignment);
		this._alignment = alignment;
	}

	@Override
	public VerticalAlignment getVerticalAlignment() {
		return _verticalAlignment;
	}

	@Override
	public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
		Validations.argNotNull(verticalAlignment);
		this._verticalAlignment = verticalAlignment;
	}

	@Override
	public boolean isWrapText() {
		return _wrapText;
	}

	@Override
	public void setWrapText(boolean wrapText) {
		this._wrapText = wrapText;
	}

	@Override
	public BorderType getBorderLeft() {
		return _borderLeft;
	}

	@Override
	public void setBorderLeft(BorderType borderLeft) {
		Validations.argNotNull(borderLeft);
		this._borderLeft = borderLeft;
	}

	@Override
	public BorderType getBorderTop() {
		return _borderTop;
	}

	@Override
	public void setBorderTop(BorderType borderTop) {
		Validations.argNotNull(borderTop);
		this._borderTop = borderTop;
	}

	@Override
	public BorderType getBorderRight() {
		return _borderRight;
	}

	@Override
	public void setBorderRight(BorderType borderRight) {
		Validations.argNotNull(borderRight);
		this._borderRight = borderRight;
	}

	@Override
	public BorderType getBorderBottom() {
		return _borderBottom;
	}

	@Override
	public void setBorderBottom(BorderType borderBottom){
		Validations.argNotNull(borderBottom);
		this._borderBottom = borderBottom;
	}

	@Override
	public SColor getBorderTopColor() {
		return _borderTopColor;
	}

	@Override
	public void setBorderTopColor(SColor borderTopColor) {
		Validations.argNotNull(borderTopColor);
		this._borderTopColor = borderTopColor;
	}

	@Override
	public SColor getBorderLeftColor() {
		return _borderLeftColor;
	}

	@Override
	public void setBorderLeftColor(SColor borderLeftColor) {
		Validations.argNotNull(borderLeftColor);
		this._borderLeftColor = borderLeftColor;
	}

	@Override
	public SColor getBorderBottomColor() {
		return _borderBottomColor;
	}

	@Override
	public void setBorderBottomColor(SColor borderBottomColor) {
		Validations.argNotNull(borderBottomColor);
		this._borderBottomColor = borderBottomColor;
	}

	@Override
	public SColor getBorderRightColor() {
		return _borderRightColor;
	}

	@Override
	public void setBorderRightColor(SColor borderRightColor) {
		Validations.argNotNull(borderRightColor);
		this._borderRightColor = borderRightColor;
	}

	@Override
	public String getDataFormat() {
		return _dataFormat;
	}
	
	@Override
	public boolean isDirectDataFormat(){
		return _directFormat;
	}

	@Override
	public void setDataFormat(String dataFormat) {
		//set to general if null to compatible with 3.0
		if(dataFormat==null || "".equals(dataFormat.trim())){
			dataFormat = FORMAT_GENERAL;
		}
		this._dataFormat = dataFormat;
		_directFormat = false;
	}
	
	@Override
	public void setDirectDataFormat(String dataFormat){
		setDataFormat(dataFormat);
		_directFormat = true;
	}

	@Override
	public boolean isLocked() {
		return _locked;
	}

	@Override
	public void setLocked(boolean locked) {
		this._locked = locked;
	}

	@Override
	public boolean isHidden() {
		return _hidden;
	}

	@Override
	public void setHidden(boolean hidden) {
		this._hidden = hidden;
	}

	@Override
	public void copyFrom(SCellStyle src) {
		if (src == this)
			return;
		Validations.argInstance(src, CellStyleImpl.class);
		setFont(src.getFont());//assign directly
		
		setFillColor(src.getFillColor());
		setFillPattern(src.getFillPattern());
		setAlignment(src.getAlignment());
		setVerticalAlignment(src.getVerticalAlignment());
		setWrapText(src.isWrapText());

		setBorderLeft(src.getBorderLeft());
		setBorderTop(src.getBorderTop());
		setBorderRight(src.getBorderRight());
		setBorderBottom(src.getBorderBottom());
		setBorderTopColor(src.getBorderTopColor());
		setBorderLeftColor(src.getBorderLeftColor());
		setBorderBottomColor(src.getBorderBottomColor());
		setBorderRightColor(src.getBorderRightColor());

		setDataFormat(src.getDataFormat());
		setLocked(src.isLocked());
		setHidden(src.isHidden());
	}
	
	@Override
	String getStyleKey() {
		StringBuilder sb = new StringBuilder();
		sb.append(_font.getStyleKey())
		.append(".").append(_fillPattern.ordinal())
		.append(".").append(_fillColor.getHtmlColor())
		.append(".").append(_alignment.ordinal())
		.append(".").append(_verticalAlignment.ordinal())
		.append(".").append(_wrapText?"T":"F")
		.append(".").append(_borderLeft.ordinal())
		.append(".").append(_borderLeftColor.getHtmlColor())
		.append(".").append(_borderRight.ordinal())
		.append(".").append(_borderRightColor.getHtmlColor())
		.append(".").append(_borderTop.ordinal())
		.append(".").append(_borderTopColor.getHtmlColor())
		.append(".").append(_borderBottom.ordinal())
		.append(".").append(_borderBottomColor.getHtmlColor())
		.append(".").append(_dataFormat)
		.append(".").append(_locked?"T":"F")
		.append(".").append(_hidden?"T":"F");
		return sb.toString();
	}

	@Override
	public void setBorderLeft(BorderType borderLeft, SColor color) {
		setBorderLeft(borderLeft);
		setBorderLeftColor(color);
	}

	@Override
	public void setBorderTop(BorderType borderTop, SColor color) {
		setBorderTop(borderTop);
		setBorderTopColor(color);
	}

	@Override
	public void setBorderRight(BorderType borderRight, SColor color) {
		setBorderRight(borderRight);
		setBorderRightColor(color);
	}

	@Override
	public void setBorderBottom(BorderType borderBottom, SColor color) {
		setBorderBottom(borderBottom);
		setBorderBottomColor(color);
	}

	//ZSS-780
	@Override
	public SColor getBackColor() {
		return _backColor;
	}

	//ZSS-780
	@Override
	public void setBackgroundColor(SColor backColor) {
		Validations.argNotNull(backColor);
		this._backColor = backColor;
		_patternHtml = null; //clear cache
	}

	//ZSS-841
	@Override
	public String getFillPatternHtml() {
		final FillPattern pattern = getFillPattern();
		if (pattern == FillPattern.NONE || pattern == FillPattern.SOLID) {
			return "";
		}
		if (_patternHtml == null) {
			_patternHtml = getFillPatternHtml(this);
		}
		return _patternHtml;
	}

	//ZSS-841
	private static byte[][] _PATTERN_BYTES = {
		null, //NO_FILL
		
		// 00000000
		// 00000000
		// 00000000
		// 00000000
		new byte[] {(byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00}, //SOLID, //SOLID_FOREGROUND
		
		// 01010101
		// 10101010
		// 01010101
		// 10101010
		new byte[] {(byte) 0x55, (byte) 0xaa, (byte) 0x55, (byte) 0xaa}, //MEDIUM_GRAY, //FINE_DOTS
		
		// 11101110
		// 10111011
		// 11101110
		// 10111011
		new byte[] {(byte) 0xee, (byte) 0xbb, (byte) 0xee, (byte) 0xbb}, //DARK_GRAY, //ALT_BARS
		
		// 10001000
		// 00100010
		// 10001000
		// 00100010
		new byte[] {(byte) 0x88, (byte) 0x22, (byte) 0x88, (byte) 0x22}, //LIGHT_GRAY, //SPARSE_DOTS
		
		// 11111111
		// 00000000
		// 00000000
		// 11111111
		new byte[] {(byte) 0xff, (byte) 0x00, (byte) 0x00, (byte) 0xff}, //DARK_HORIZONTAL, //THICK_HORZ_BANDS
		
		// 00110011
		// 00110011
		// 00110011
		// 00110011
		new byte[] {(byte) 0x33, (byte) 0x33, (byte) 0x33, (byte) 0x33}, //DARK_VERTICAL, //THICK_VERT_BANDS
		
		// 10011001
		// 11001100
		// 01100110
		// 00110011
		new byte[] {(byte) 0x99, (byte) 0xcc, (byte) 0x66, (byte) 0x33}, //DARK_DOWN, //THICK_BACKWARD_DIAG
		
		// 10011001
		// 00110011
		// 01100110
		// 11001100
		new byte[] {(byte) 0x99, (byte) 0x33, (byte) 0x66, (byte) 0xcc}, //DARK_UP, //THICK_FORWARD_DIAG
		
		// 10011001
		// 10011001
		// 01100110
		// 01100110
		new byte[] {(byte) 0x99, (byte) 0x99, (byte) 0x66, (byte) 0x66}, //DARK_GRID, //BIG_SPOTS
		
		// 10011001
		// 11111111
		// 01100110
		// 11111111
		new byte[] {(byte) 0x99, (byte) 0xff, (byte) 0x66, (byte) 0xff}, //DARK_TRELLIS, //BRICKS
		
		// 00000000
		// 00000000
		// 00000000
		// 11111111
		new byte[] {(byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0xff}, //LIGHT_HORIZONTAL, //THIN_HORZ_BANDS
		
		// 00100010
		// 00100010
		// 00100010
		// 00100010
		new byte[] {(byte) 0x22, (byte) 0x22, (byte) 0x22, (byte) 0x22}, //LIGHT_VERTICAL, //THIN_VERT_BANDS
		
		// 00010001
		// 10001000
		// 01000100
		// 00100010
		new byte[] {(byte) 0x11, (byte) 0x88, (byte) 0x44, (byte) 0x22}, //LIGHT_DOWN, //THIN_BACKWARD_DIAG
		
		// 10001000
		// 00010001
		// 00100010
		// 01000100
		new byte[] {(byte) 0x88, (byte) 0x11, (byte) 0x22, (byte) 0x44}, //LIGHT_UP, //THIN_FORWARD_DIAG
		
		// 00100010
		// 00100010
		// 00100010
		// 11111111 
		new byte[] {(byte) 0x22, (byte) 0x22, (byte) 0x22, (byte) 0xff}, //LIGHT_GRID, //SQUARES
		
		// 01010101
		// 10001000
		// 01010101
		// 00100010
		new byte[] {(byte) 0x55, (byte) 0x88, (byte) 0x55, (byte) 0x22}, //LIGHT_TRELLIS, //DIAMONDS
		
		// 00000000
		// 10001000
		// 00000000
		// 00100010
		new byte[] {(byte) 0x00, (byte) 0x88, (byte) 0x00, (byte) 0x22}, //GRAY125, //LESS_DOTS 
		
		// 00000000
		// 00100000
		// 00000000
		// 00000010
		new byte[] {(byte) 0x00, (byte) 0x20, (byte) 0x00, (byte) 0x02}, //GRAY0625 //LEAST_DOTS
	};
	
	//ZSS-841
	public static String getFillPatternHtml(SCellStyle style) {
		BufferedImage image = new BufferedImage(8, 4, BufferedImage.TYPE_INT_ARGB);
		Graphics2D g2 = image.createGraphics();
		// background color
		byte[] rgb = style.getBackColor().getRGB();
		g2.setColor(new Color(((int)rgb[0]) & 0xff , ((int)rgb[1]) & 0xff, ((int)rgb[2]) & 0xff));
		g2.fillRect(0, 0, 8, 4);
		// foreground color
		rgb = style.getFillColor().getRGB();
		g2.setColor(new Color(((int)rgb[0]) & 0xff , ((int)rgb[1]) & 0xff, ((int)rgb[2]) & 0xff));
		byte[] patb = _PATTERN_BYTES[style.getFillPattern().ordinal()];
		for (int y = 0; y < 4; ++y) {
			byte b = patb[y];
			if (b == 0) continue; // all zero case
			if (b == 0xff) {
				g2.drawLine(0, y, 7, y);
				continue;
			}
			for (int x = 0; x < 8; ++x) {
				if ((b & 0x80) != 0) {
					g2.drawLine(x, y, x, y);
				}
				b <<= 1;
			}
		}
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try {
			ImageIO.write(image, "png", os);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				os.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		StringBuilder sb = new StringBuilder();
		sb.append("background-image:url(data:image/png;base64,");
		String base64 = Base64.encodeBase64String(os.toByteArray());
		sb.append(base64).append(");");
		return sb.toString();
	}
}
