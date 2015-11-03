/* ConditonalFormattingRuleImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 10:16:50 AM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.List;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.zkoss.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.zkoss.zss.model.SColorScale;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.SExtraStyle;
import org.zkoss.zss.model.SDataBar;
import org.zkoss.zss.model.SIconSet;
/**
 * @author henri
 * @since 3.8.2
 * 
 */
//ZSS-1138
public class ConditionalFormattingRuleImpl implements SConditionalFormattingRule {
	private RuleType type;
	private RuleOperator operator;
	private Integer priority;
	private SExtraStyle style;
	private boolean stopIfTrue;
	private List<String> formulas;
	private RuleTimePeriod timePeriod;
	private Long rank;
	private boolean percent;
	private boolean bottom;
	private SColorScale colorScale;
	private SDataBar dataBar;
	private SIconSet iconSet;
	private String text;
	private boolean notAboveAverage;
	private boolean equalAverage;
	private Integer standardDeviation; //1 ~ 3
	
	@Override
	public RuleType getType() {
		return type;
	}

	public void setType(RuleType type) {
		this.type = type;
	}
	
	@Override
	public RuleOperator getOperator() {
		return operator;
	}
	
	public void setOperator(RuleOperator operator) {
		this.operator = operator; 
	}

	@Override
	public Integer getPriority() {
		return priority;
	}
	
	public void setPriority(int priority) {
		this.priority = priority;
	}

	@Override
	public SExtraStyle getExtraStyle() {
		return style;
	}
	
	public void setExtraStyle(SExtraStyle style) {
		this.style = style;
	}

	@Override
	public boolean isStopIfTrue() {
		return stopIfTrue;
	}
	
	public void setStopIfTrue(boolean b) {
		stopIfTrue = b;
	}

	@Override
	public List<String> getFormulas() {
		return formulas;
	}
	
	public void addFormula(String formula) {
		if (formulas == null) {
			formulas = new ArrayList<String>();
		}
		formulas.add(formula);
	}

	@Override
	public RuleTimePeriod getTimePeriod() {
		return timePeriod;
	}
	
	public void setTimePeriod(RuleTimePeriod timePeriod) {
		this.timePeriod = timePeriod;
	}

	@Override
	public Long getRank() {
		return rank;
	}
	
	public void setRank(long rank) {
		this.rank = rank;
	}

	@Override
	public boolean isPercent() {
		return percent;
	}
	
	public void setPercent(boolean b) {
		percent = b;
	}

	@Override
	public boolean isBottom() {
		return bottom;
	}

	public void setBottom(boolean b) {
		bottom = b;
	}
	
	@Override
	public SColorScale getColorScale() {
		return colorScale;
	}
	
	public void setColorScale(SColorScale colorScale) {
		this.colorScale = colorScale;
	}

	@Override
	public SDataBar getDataBar() {
		return dataBar;
	}

	public void setDataBar(SDataBar dataBar) {
		this.dataBar = dataBar;
	}
	
	@Override
	public SIconSet getIconSet() {
		return iconSet;
	}

	public void setIconSet(SIconSet iconSet) {
		this.iconSet = iconSet;
	}
	
	@Override
	public String getText() {
		return text;
	}
	
	public void setText(String text) {
		this.text = text;
	}

	@Override
	public boolean isAboveAverage() {
		return !notAboveAverage;
	}
	
	public void setAboveAverage(boolean b) {
		notAboveAverage = !b;
	}

	@Override
	public boolean isEqualAverage() {
		return equalAverage;
	}
	
	public void setEqualAverage(boolean b) {
		equalAverage = b;
	}

	@Override
	public Integer getStandardDeviation() {
		return standardDeviation;
	}
	
	public void setStandardDeviation(int s) {
		standardDeviation = s;
	}

}
