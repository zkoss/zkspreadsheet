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
package org.zkoss.zss.ngmodel;

import java.util.List;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
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
