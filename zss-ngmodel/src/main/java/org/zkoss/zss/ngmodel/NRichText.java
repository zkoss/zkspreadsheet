package org.zkoss.zss.ngmodel;

import java.util.List;

public interface NRichText {

	public interface Segment {
		public String getText();	
		public NFont getFont();
	}
	
	public String getText();
	
	public List<Segment> getSegments();
	
	public void addSegment(String text,NFont font);
	
	public void clearSegments();
	
}
