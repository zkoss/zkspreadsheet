package org.zkoss.zss.api.model;


public interface NCellStyle {

	public enum FillPattern {
		NO_FILL, SOLID_FOREGROUND, FINE_DOTS, ALT_BARS, SPARSE_DOTS, THICK_HORZ_BANDS, THICK_VERT_BANDS, THICK_BACKWARD_DIAG, THICK_FORWARD_DIAG, BIG_SPOTS, BRICKS, THIN_HORZ_BANDS, THIN_VERT_BANDS, THIN_BACKWARD_DIAG, THIN_FORWARD_DIAG, SQUARES, DIAMONDS, LESS_DOTS, LEAST_DOTS
	}

	public enum Alignment {
		GENERAL, LEFT, CENTER, RIGHT, FILL, JUSTIFY, CENTER_SELECTION
	}

	public enum VerticalAlignment {
		TOP, CENTER, BOTTOM, JUSTIFY
	}

	public enum BorderType {
		NONE, THIN, MEDIUM, DASHED, HAIR, THICK, DOUBLE, DOTTED, MEDIUM_DASHED, DASH_DOT, MEDIUM_DASH_DOT, DASH_DOT_DOT, MEDIUM_DASH_DOT_DOT, SLANTED_DASH_DOT;
	}

	public NFont getFont();

	public void setFont(NFont nfont);

	public void cloneAttribute(NCellStyle src);

	public NColor getBackgroundColor();

	public void setBackgroundColor(NColor color);

	public FillPattern getFillPattern();

	public void setFillPattern(FillPattern pattern);

	public void setAlignment(Alignment alignment);

	public Alignment getAlignment();

	public void setVerticalAlignment(VerticalAlignment alignment);

	public VerticalAlignment getVerticalAlignment();

	public boolean isWrapText();

	public void setWrapText(boolean wraptext);

	public void setBorderLeft(BorderType borderType);

	public BorderType getBorderLeft();

	public void setBorderTop(BorderType borderType);

	public BorderType getBorderTop();

	public void setBorderRight(BorderType borderType);

	public BorderType getBorderRight();

	public void setBorderBottom(BorderType borderType);

	public BorderType getBorderBottom();

	public void setBorderTopColor(NColor color);

	public NColor getBorderTopColor();

	public void setBorderLeftColor(NColor color);

	public NColor getBorderLeftColor();

	public void setBorderBottomColor(NColor color);

	public NColor getBorderBottomColor();

	public void setBorderRightColor(NColor color);

	public NColor getBorderRightColor();
}
