package org.zkoss.zss.ngmodel.impl;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.NRichText;

public class ReadOnlyRichTextImpl extends AbstractRichTextAdv {

	private static final long serialVersionUID = 1L;
	private List<Segment> segments;
	
	private NRichText richText;

	public ReadOnlyRichTextImpl(String text, NFont font) {
		segments = new ArrayList<Segment>(1);
		segments.add(new SegmentImpl(text,font));
	}
	public ReadOnlyRichTextImpl(NRichText richText) {
		this.richText = richText;
	}

	@Override
	public String getText() {
		return richText==null?segments.get(0).getText():richText.getText();
	}

	@Override
	public NFont getFont() {
		return richText==null?segments.get(0).getFont():richText.getFont();
	}

	@SuppressWarnings("unchecked")
	@Override
	public List<Segment> getSegments() {
		return richText==null?Collections.unmodifiableList(segments):richText.getSegments();
	}

	@Override
	public void addSegment(String text, NFont font) {
		throw new UnsupportedOperationException("readonly rich text");
	}

	@Override
	public void clearSegments() {
		throw new UnsupportedOperationException("readonly rich text");
	}
	@Override
	public AbstractRichTextAdv clone() {
		return richText==null?new ReadOnlyRichTextImpl(segments.get(0).getText(),segments.get(0).getFont()):new ReadOnlyRichTextImpl(richText);
	}

}
