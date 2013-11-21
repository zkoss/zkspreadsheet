package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.util.Validations;

public class CellStyleImpl extends AbstractCellStyle{

	public static final String COLOR_WHITE = "#FFFFFF";
	public static final String COLOR_BLACK = "#000000";
	public static final String FORMAT_GENERAL = "General";

	private String fontName = "Arial";
	private String fontColor = COLOR_BLACK;
	private FontBoldweight fontBoldweight = FontBoldweight.NORMAL;
	private int fontHeight = 11;
	private boolean fontItalic = false;
	private boolean fontStrikeout = false;
	private FontTypeOffset fontTypeOffset = FontTypeOffset.NONE;
	private FontUnderline fontUnderline = FontUnderline.NONE;

	private String backgroundColor = COLOR_WHITE;
	private FillPattern fillPattern = FillPattern.NO_FILL;
	private Alignment alignment = Alignment.LEFT;
	private VerticalAlignment verticalAlignment = VerticalAlignment.BOTTOM;
	private boolean wrapText = false;

	private BorderType borderLeft = BorderType.NONE;
	private BorderType borderTop = BorderType.NONE;
	private BorderType borderRight = BorderType.NONE;
	private BorderType borderBottom = BorderType.NONE;
	private String borderTopColor = COLOR_BLACK;
	private String borderLeftColor = COLOR_BLACK;
	private String borderBottomColor = COLOR_BLACK;
	private String borderRightColor = COLOR_BLACK;

	private String dataFormat = FORMAT_GENERAL;
	private boolean locked = true;// default locked as excel.
	private boolean hidden = false;

	public String getFontName() {
		return fontName;
	}

	public void setFontName(String fontName) {
		this.fontName = fontName;
	}

	public String getFontColor() {
		return fontColor;
	}

	public void setFontColor(String fontColor) {
		this.fontColor = fontColor;
	}

	public FontBoldweight getFontBoldweight() {
		return fontBoldweight;
	}

	public void setFontBoldweight(FontBoldweight fontBoldweight) {
		this.fontBoldweight = fontBoldweight;
	}

	public int getFontHeight() {
		return fontHeight;
	}

	public void setFontHeight(int fontHeight) {
		this.fontHeight = fontHeight;
	}

	public boolean isFontItalic() {
		return fontItalic;
	}

	public void setFontItalic(boolean fontItalic) {
		this.fontItalic = fontItalic;
	}

	public boolean isFontStrikeout() {
		return fontStrikeout;
	}

	public void setFontStrikeout(boolean fontStrikeout) {
		this.fontStrikeout = fontStrikeout;
	}

	public FontTypeOffset getFontTypeOffset() {
		return fontTypeOffset;
	}

	public void setFontTypeOffset(FontTypeOffset fontTypeOffset) {
		this.fontTypeOffset = fontTypeOffset;
	}

	public FontUnderline getFontUnderline() {
		return fontUnderline;
	}

	public void setFontUnderline(FontUnderline fontUnderline) {
		this.fontUnderline = fontUnderline;
	}

	public String getBackgroundColor() {
		return backgroundColor;
	}

	public void setBackgroundColor(String backgroundColor) {
		this.backgroundColor = backgroundColor;
	}

	public FillPattern getFillPattern() {
		return fillPattern;
	}

	public void setFillPattern(FillPattern fillPattern) {
		this.fillPattern = fillPattern;
	}

	public Alignment getAlignment() {
		return alignment;
	}

	public void setAlignment(Alignment alignment) {
		this.alignment = alignment;
	}

	public VerticalAlignment getVerticalAlignment() {
		return verticalAlignment;
	}

	public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
		this.verticalAlignment = verticalAlignment;
	}

	public boolean isWrapText() {
		return wrapText;
	}

	public void setWrapText(boolean wrapText) {
		this.wrapText = wrapText;
	}

	public BorderType getBorderLeft() {
		return borderLeft;
	}

	public void setBorderLeft(BorderType borderLeft) {
		this.borderLeft = borderLeft;
	}

	public BorderType getBorderTop() {
		return borderTop;
	}

	public void setBorderTop(BorderType borderTop) {
		this.borderTop = borderTop;
	}

	public BorderType getBorderRight() {
		return borderRight;
	}

	public void setBorderRight(BorderType borderRight) {
		this.borderRight = borderRight;
	}

	public BorderType getBorderBottom() {
		return borderBottom;
	}

	public void setBorderBottom(BorderType borderBottom) {
		this.borderBottom = borderBottom;
	}

	public String getBorderTopColor() {
		return borderTopColor;
	}

	public void setBorderTopColor(String borderTopColor) {
		this.borderTopColor = borderTopColor;
	}

	public String getBorderLeftColor() {
		return borderLeftColor;
	}

	public void setBorderLeftColor(String borderLeftColor) {
		this.borderLeftColor = borderLeftColor;
	}

	public String getBorderBottomColor() {
		return borderBottomColor;
	}

	public void setBorderBottomColor(String borderBottomColor) {
		this.borderBottomColor = borderBottomColor;
	}

	public String getBorderRightColor() {
		return borderRightColor;
	}

	public void setBorderRightColor(String borderRightColor) {
		this.borderRightColor = borderRightColor;
	}

	public String getDataFormat() {
		return dataFormat;
	}

	public void setDataFormat(String dataFormat) {
		this.dataFormat = dataFormat;
	}

	public boolean isLocked() {
		return locked;
	}

	public void setLocked(boolean locked) {
		this.locked = locked;
	}

	public boolean isHidden() {
		return hidden;
	}

	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	public void copyTo(AbstractCellStyle dest) {
		
		Validations.argInstance(dest, CellStyleImpl.class);
		if (dest == this)
			return;
		CellStyleImpl another = (CellStyleImpl)dest;
		another.fontName = fontName;
		another.fontColor = fontColor;
		another.fontBoldweight = fontBoldweight;
		another.fontHeight = fontHeight;
		another.fontItalic = fontItalic;
		another.fontStrikeout = fontStrikeout;
		another.fontTypeOffset = fontTypeOffset;
		another.fontUnderline = fontUnderline;

		another.backgroundColor = backgroundColor;
		another.fillPattern = fillPattern;
		another.alignment = alignment;
		another.verticalAlignment = verticalAlignment;
		another.wrapText = wrapText;

		another.borderLeft = borderLeft;
		another.borderTop = borderTop;
		another.borderRight = borderRight;
		another.borderBottom = borderBottom;
		another.borderTopColor = borderTopColor;
		another.borderLeftColor = borderLeftColor;
		another.borderBottomColor = borderBottomColor;
		another.borderRightColor = borderRightColor;

		another.dataFormat = dataFormat;
		another.locked = locked;
		another.hidden = hidden;
	}
}
