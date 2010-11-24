/* SortLevel.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Aug 17, 2010 6:44:08 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.sort;

/**
 * @author Sam
 *
 */
public class SortLevel {

	/**
	 * Indicate sort target, row or column index number
	 */
	public int sortIndex = -1;
	
	/**
	 * Indicate sort target content is number or value
	 * <p> Default false, sort by value
	 */
	public boolean sortNumber;
	
	/**
	 * Indicate sort order, either descending or ascending
	 * <p> Default false: ascending
	 */
	public boolean order;

	/**
	 * see numeric String as number or not
	 */
	public int dataOption;
	
	public final static boolean ASCENDING = false;
	public final static boolean DESCENDING = true;
}
