/* SConditionalFormattingRule.java

	Purpose:
		
	Description:
		
	History:
		Oct 23, 2015 12:01:39 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTIconSet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STTimePeriod;

/**
 * @author henri
 * @since 3.8.2
 */
public interface SConditionalFormattingRule {
	/** Returns type of this rule */
	RuleType getType();
	
	/** Returns the comparison operator if needed */
	RuleOperator getOperator();
	
	/** Returns the applying priority if convered region overlapped */
	int getPriority();
	
	/** Returns the applied color if match */
	SExtraStyle getExtraStyle();
	
	/** Returns whether stop if true */
	boolean isStopIfTrue();  // default false
	
	/** Returns the associated formulas (used with "cellIs" type) */
	List<String> getFormulas();
	
	/** Returns the timePeriod operator (used with "timePeriod" type)*/
	RuleTimePeriod getTimePeriod();

	/** Returns the rank used with "top10" type */
	long getRank();
	
	/** Returns whether a percentage (used with "top10" type) */
	boolean isPercent();

	/** Returns whether bottom value (used with "top10" type).
	 * false means check "above 10 ...";
	 * true means check "below 10 ...". 
	 */
	boolean isBottom();

	/** Returns the color scale (used with "colorScale" type) */
	SColorScale getColorScale();
	
	/** Returns the data bar (used with "dataBar" type) */
	SDataBar getDataBar();
	
	/** Returns the icon set (used with "iconSet" type) */
	SIconSet getIconSet();
	
	/** Returns the text for comparison (used with "containsText" type).
	 *  {@link #getFormula} will return the  formula for doing the search operation.
	 */
	String getText();

	/** Returns whether above average (used with "aboveAverage" type).
	 * false means "check below average ..."; 
	 * true means "check above average ..." (default true)
	 */
	boolean isAboveAverage();
	
	/** Returns whether equal average (used with "aboveAverage" type). */
	boolean isEqualAverage();

	/** Returns 1st/2nd/3rd standard deviation average (used with "aboveAverage" type */
	int getStandardDeviation();


	public enum RuleTimePeriod {
		TODAY("today", 1),       
		YESTERDAY("yesterday", 2),   
		TOMORROW("Tomorrow", 3),    
		LAST_7_DAYS("last7Days", 4), 
		THIS_MONTH("thisMonth", 5),  
		LAST_MONTH("lastMonth", 6),  
		NEXT_MONTH("nextMonth", 7),  
		THIS_WEEK("thisWeek", 8),   
		LAST_WEEK("lastWeek", 9),   
		NEXT_WEEK("nextWeek", 10);  
	
		public final String name;
		public final int value;
		
		RuleTimePeriod(String name, int value) {
			this.name = name;
			this.value = value;
		}
	}
	
	public enum RuleType {
		EXPRESSION("expression", 1),
		CELL_IS("cellIs", 2),
		COLOR_SCALE("colorScale", 3),        
		DATA_BAR("dataBar", 4),           
		ICON_SET("iconSet", 5),           
		TOP_10("top10", 6),             
		UNIQUE_VALUES("uniqueValues", 7),      
		DUPLICATE_VALUES("duplicateValues", 8),   
		CONTAINS_TEXT("containsText", 9),      
		NOT_CONTAINS_TEXT("notContainsText", 10), 
		BEGINS_WITH("beginsWith", 11),       
		ENDS_WITH("endsWith", 12),         
		CONTAINS_BLANKS("containsBlanks", 13),   
		NOT_CONTAINS_BLANKS("notContainsBlanks", 14),
		CONTAINS_ERRORS("containsErrors", 15),   
		NOT_CONTAINS_ERRORS("containsErrors", 16),
		TIME_PERIOD("timePeriod", 17),       
		ABOVE_AVERAGE("aboveAverage", 18);     
		
		public final String name;
		public final int value;
		
		RuleType(String name, int value) {
			this.name = name;
			this.value = value;
		}
	}
	
	public enum RuleOperator {
		LESS_THAN("lessThan", 1),
		LESS_THAN_OR_EQUAL("lessThanOrEqual", 2),
		EQUAL("equal" , 3),
		NOT_EQUAL("notEqual" , 4),
		GREATER_THAN_OR_EQUAL("greaterThanOrEqual" , 5),
		GREATER_THAN("greaterThan" , 6),
		BETWEEN("between" , 7), /* the two values is in formula list */
		NOT_BETWEEN("notBetween" , 8),
		CONTAINS_TEXT("containsText" , 9),
		NOT_CONTAINS("notContains" , 10),
		BEGINS_WITH("beginsWith" , 11),
		ENDS_WITH("endsWith" , 12);
	
		public final String name;
		public final int value;
		
		RuleOperator(String name, int value) {
			this.name = name;
			this.value = value;
		}
	}
}