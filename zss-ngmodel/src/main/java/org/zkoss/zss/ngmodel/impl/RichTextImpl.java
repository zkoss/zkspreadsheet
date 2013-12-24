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
package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;

import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class RichTextImpl extends AbstractRichTextAdv {
	private static final long serialVersionUID = 1L;

	List<SegmentImpl> segments = new LinkedList<SegmentImpl>();


	@Override
	public String getText() {
		if (segments.size() == 0) {
			return "";
		}
		StringBuilder sb = new StringBuilder();
		for (Segment s : segments) {
			sb.append(s.getText());
		}
		return sb.toString();
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public List<Segment> getSegments() {
		return Collections.unmodifiableList((List) segments);
	}

	@Override
	public void addSegment(String text, NFont font) {
		Validations.argNotNull(text,font);
		if("".equals(text)) return;
		segments.add(new SegmentImpl(text, font));
	}

	@Override
	public void clearSegments() {
		segments.clear();
	}

	@Override
	public NFont getFont() {
		if (segments.size() == 0) {
			return null;
		}
		return segments.get(0).getFont();
	}

}
