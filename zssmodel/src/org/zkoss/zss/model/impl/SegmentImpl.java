package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.SRichText.Segment;

class SegmentImpl implements Segment, Serializable {
	private static final long serialVersionUID = 1L;
	private String text;
	private SFont font;

	SegmentImpl(String text, SFont font) {
		this.text = text;
		this.font = font;
	}

	@Override
	public String getText() {
		return text;
	}

	@Override
	public SFont getFont() {
		return font;
	}

}