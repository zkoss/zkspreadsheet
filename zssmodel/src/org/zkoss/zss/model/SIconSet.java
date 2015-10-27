/* SIconSet.java

	Purpose:
		
	Description:
		
	History:
		Oct 23, 2015 3:57:16 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import java.util.List;

/**
 * @author henri
 *
 */
public interface SIconSet {
	/** Returns the icon set name */
	IconSetType getType();
	
	/** Returns the value objects */
	List<SCFValueObject> getCFValueObjects(); 

	/** Returns whether in percent */
	boolean isPercent();
	
	/** Returns whether reverse the order */
	boolean isReverse();
	
	/** Returns whether show the value */
	boolean isShowValue();
	
	public enum IconSetType {
		X_3_ARROWS("3Arrows", 1),
		X_3_ARROWS_GRAY("3ArrowsGray", 2),
		X_3_FLAGS("3Flags", 3),
		X_3_TRAFFIC_LIGHTS_1("3TrafficLights1", 4),
		X_3_TRAFFIC_LIGHTS_2("3TrafficLights2", 5),
		X_3_SIGNS("3Signs", 6),
		X_3_SYMBOLS("3Symbols", 7),
		X_3_SYMBOLS_2("3Symbols2", 8),
		X_4_ARROWS("4Arrows", 9),
		X_4_ARROWS_GRAY("4ArrowsGray", 10),
		X_4_RED_TO_BLACK("4RedToBlack", 11),
		X_4_RATING("4Rating", 12),
		X_4_TRAFFIC_LIGHTS("4TrafficLights", 13),
		X_5_ARROWS("5Arrows", 14),
		X_5_ARROWS_GRAY("5ArrowsGray", 15),
		X_5_RATING("5Rating", 16),
		X_5_QUARTERS("5Quarters", 17);
	
		public final String name;
	   	public final int value;	    		   	IconSetType(String name, int value) {	  		this.name = name;	   		this.value = value;	   	}    			}
	
	}    