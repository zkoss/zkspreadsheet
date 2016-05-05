/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.SRichText;
import org.zkoss.zss.model.SRichText.Segment;
/**
 * 
 * @author Dennis
 * @since 3.5.0
 */
public class ReadOnlyRichTextImpl extends AbstractRichTextAdv {

	private static final long serialVersionUID = 1L;
	private List<Segment> _segments;
	
	private SRichText _richText;

	public ReadOnlyRichTextImpl(String text, SFont font) {
		_segments = new ArrayList<Segment>(1);
		_segments.add(new SegmentImpl(text,font));
	}
	public ReadOnlyRichTextImpl(SRichText richText) {
		this._richText = richText;
	}

	@Override
	public String getText() {
		return _richText==null?_segments.get(0).getText():_richText.getText();
	}

	@Override
	public SFont getFont() {
		return _richText==null?_segments.get(0).getFont():_richText.getFont();
	}

	@SuppressWarnings("unchecked")
	@Override
	public List<Segment> getSegments() {
		return _richText==null?Collections.unmodifiableList(_segments):_richText.getSegments();
	}

	@Override
	public void addSegment(String text, SFont font) {
		throw new UnsupportedOperationException("readonly rich text");
	}

	@Override
	public void clearSegments() {
		throw new UnsupportedOperationException("readonly rich text");
	}
	
	@Override
	public AbstractRichTextAdv clone() {
		return _richText==null?new ReadOnlyRichTextImpl(_segments.get(0).getText(),_segments.get(0).getFont()):new ReadOnlyRichTextImpl(_richText);
	}
	
	@Override
	public int getHeightPoints() {
		int highest = 0;
		for (Segment ss : getSegments()) {
			int p = ss.getFont().getHeightPoints();
			if(p > highest) 
				highest = p;
		}
		return highest;
	}
	
	//ZSS-1183
	@Override
	/*package*/ AbstractRichTextAdv cloneRichText(SBook book) {
		if (_richText == null) {
			final SFont font0 = _segments.get(0).getFont();
			final SFont font = font0 == null ? null :
				((AbstractFontAdv)font0).cloneFont(book);
			return new ReadOnlyRichTextImpl(_segments.get(0).getText(), font);
		} else {
			AbstractRichTextAdv richText = ((AbstractRichTextAdv)_richText).cloneRichText(book);
			return new ReadOnlyRichTextImpl(richText);
		}
	}

}
