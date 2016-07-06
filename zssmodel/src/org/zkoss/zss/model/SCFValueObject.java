/* SCFValueObject.java

	Purpose:
		
	Description:
		
	History:
		Oct 23, 2015 4:06:27 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

/**
 * Conditional Formatting value object.
 * @author henri
 *
 */
public interface SCFValueObject {
	/** Return the type of this conditional formatting value object */
	CFValueObjectType getType(); 
	
	/** Return the value of this condition formatting value object */
	String getValue();
	
	/** Whether greaterOrEqual(true) or greater(false) */
	boolean isGreaterOrEqual();

	public enum CFValueObjectType {
		NUM("num", 1),
		PERCENT("percent", 2),
		MAX("max", 3),
		MIN("min", 4),
		FORMULA("formula", 5),
		PERCENTILE("percentile", 6);
	
		public final String name;
	   	public final int value;
	    	
	   	CFValueObjectType(String name, int value) {
	  		this.name = name;
	   		this.value = value;
	   	}    		
	}
}