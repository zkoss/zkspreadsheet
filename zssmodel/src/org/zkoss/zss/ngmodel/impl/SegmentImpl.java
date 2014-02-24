package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.NRichText.Segment;

class SegmentImpl implements Segment, Serializable {
	private static final long serialVersionUID = 1L;
	private String text;
	private NFont font;

	SegmentImpl(String text, NFont font) {
		this.text = text;
		this.font = font;
	}

	@Override
	public String getText() {
		return text;
	}

	@Override
	public NFont getFont() {
		return font;
	}

}