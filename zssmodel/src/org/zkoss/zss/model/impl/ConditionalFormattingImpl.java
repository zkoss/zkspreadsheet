/* ConditionalFormattingImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 9:21:47 AM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.XSSFConditionalFormatting;
import org.zkoss.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SConditionalFormatting;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.SSheet;

/**
 * @author henri
 * @since 3.8.2
 */
public class ConditionalFormattingImpl implements SConditionalFormatting, Serializable {
	private static final long serialVersionUID = -2075561905182660069L;
	
	private SSheet sheet;
	private List<CellRegion> regions;
	private List<SConditionalFormattingRule> rules;
	
	public ConditionalFormattingImpl(SSheet sheet) {
		this.sheet = sheet;
	}

	@Override
	public SSheet getSheet() {
		return sheet;
	}

	/* (non-Javadoc)
	 * @see org.zkoss.zss.model.SConditionalFormatting#getRegions()
	 */
	@Override
	public List<CellRegion> getRegions() {
		return regions;
	}

	@Override
	public List<SConditionalFormattingRule> getRules() {
		return rules;
	}

	public void addRule(SConditionalFormattingRule rule) {
		if (rules == null) {
			rules = new ArrayList<SConditionalFormattingRule>();
		}
		rules.add(rule);
	}
	
	public void addRegion(CellRegion region) {
		if (regions == null) {
			regions = new ArrayList<CellRegion>();
		}
		regions.add(region);
	}
}
