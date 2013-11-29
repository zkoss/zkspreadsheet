package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;

import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.util.Validations;

public class RichTextImpl extends RichTextAdv {
	private static final long serialVersionUID = 1L;

	List<SegmentImpl> segments = new LinkedList<RichTextImpl.SegmentImpl>();

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

}
