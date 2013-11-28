package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.util.Validations;

public class CellStyleImpl extends AbstractCellStyle {
	private static final long serialVersionUID = 1L;

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

	@Override
	public String getFontName() {
		return fontName;
	}

	@Override
	public void setFontName(String fontName) {
		this.fontName = fontName;
	}

	@Override
	public String getFontColor() {
		return fontColor;
	}

	@Override
	public void setFontColor(String fontColor) {
		this.fontColor = fontColor;
	}

	@Override
	public FontBoldweight getFontBoldweight() {
		return fontBoldweight;
	}

	@Override
	public void setFontBoldweight(FontBoldweight fontBoldweight) {
		this.fontBoldweight = fontBoldweight;
	}

	@Override
	public int getFontHeight() {
		return fontHeight;
	}

	@Override
	public void setFontHeight(int fontHeight) {
		this.fontHeight = fontHeight;
	}

	@Override
	public boolean isFontItalic() {
		return fontItalic;
	}

	@Override
	public void setFontItalic(boolean fontItalic) {
		this.fontItalic = fontItalic;
	}

	@Override
	public boolean isFontStrikeout() {
		return fontStrikeout;
	}

	@Override
	public void setFontStrikeout(boolean fontStrikeout) {
		this.fontStrikeout = fontStrikeout;
	}

	@Override
	public FontTypeOffset getFontTypeOffset() {
		return fontTypeOffset;
	}

	@Override
	public void setFontTypeOffset(FontTypeOffset fontTypeOffset) {
		this.fontTypeOffset = fontTypeOffset;
	}

	@Override
	public FontUnderline getFontUnderline() {
		return fontUnderline;
	}

	@Override
	public void setFontUnderline(FontUnderline fontUnderline) {
		this.fontUnderline = fontUnderline;
	}

	@Override
	public String getBackgroundColor() {
		return backgroundColor;
	}

	@Override
	public void setBackgroundColor(String backgroundColor) {
		this.backgroundColor = backgroundColor;
	}

	@Override
	public FillPattern getFillPattern() {
		return fillPattern;
	}

	@Override
	public void setFillPattern(FillPattern fillPattern) {
		this.fillPattern = fillPattern;
	}

	@Override
	public Alignment getAlignment() {
		return alignment;
	}

	@Override
	public void setAlignment(Alignment alignment) {
		this.alignment = alignment;
	}

	@Override
	public VerticalAlignment getVerticalAlignment() {
		return verticalAlignment;
	}

	@Override
	public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
		this.verticalAlignment = verticalAlignment;
	}

	@Override
	public boolean isWrapText() {
		return wrapText;
	}

	@Override
	public void setWrapText(boolean wrapText) {
		this.wrapText = wrapText;
	}

	@Override
	public BorderType getBorderLeft() {
		return borderLeft;
	}

	@Override
	public void setBorderLeft(BorderType borderLeft) {
		this.borderLeft = borderLeft;
	}

	@Override
	public BorderType getBorderTop() {
		return borderTop;
	}

	@Override
	public void setBorderTop(BorderType borderTop) {
		this.borderTop = borderTop;
	}

	@Override
	public BorderType getBorderRight() {
		return borderRight;
	}

	@Override
	public void setBorderRight(BorderType borderRight) {
		this.borderRight = borderRight;
	}

	@Override
	public BorderType getBorderBottom() {
		return borderBottom;
	}

	@Override
	public void setBorderBottom(BorderType borderBottom) {
		this.borderBottom = borderBottom;
	}

	@Override
	public String getBorderTopColor() {
		return borderTopColor;
	}

	@Override
	public void setBorderTopColor(String borderTopColor) {
		this.borderTopColor = borderTopColor;
	}

	@Override
	public String getBorderLeftColor() {
		return borderLeftColor;
	}

	@Override
	public void setBorderLeftColor(String borderLeftColor) {
		this.borderLeftColor = borderLeftColor;
	}

	@Override
	public String getBorderBottomColor() {
		return borderBottomColor;
	}

	@Override
	public void setBorderBottomColor(String borderBottomColor) {
		this.borderBottomColor = borderBottomColor;
	}

	@Override
	public String getBorderRightColor() {
		return borderRightColor;
	}

	@Override
	public void setBorderRightColor(String borderRightColor) {
		this.borderRightColor = borderRightColor;
	}

	@Override
	public String getDataFormat() {
		return dataFormat;
	}

	@Override
	public void setDataFormat(String dataFormat) {
		this.dataFormat = dataFormat;
	}

	@Override
	public boolean isLocked() {
		return locked;
	}

	@Override
	public void setLocked(boolean locked) {
		this.locked = locked;
	}

	@Override
	public boolean isHidden() {
		return hidden;
	}

	@Override
	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	@Override
	public void copyTo(AbstractCellStyle dest) {
		if (dest == this)
			return;
		Validations.argInstance(dest, CellStyleImpl.class);
		CellStyleImpl another = (CellStyleImpl) dest;
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
