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
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.XSSFConditionalFormatting;
import org.zkoss.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SConditionalFormatting;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.impl.sys.DependencyTableAdv;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.util.Validations;

/**
 * @author henri
 * @since 3.8.2
 */
public class ConditionalFormattingImpl implements SConditionalFormatting, Serializable {
	private static final long serialVersionUID = -2075561905182660069L;
	
	private SSheet _sheet;
	private LinkedHashSet<CellRegion> _regions;
	private List<SConditionalFormattingRule> rules;
	private int _id;
	
	public ConditionalFormattingImpl(SSheet sheet) {
		this._sheet = sheet;
		this._id = ((SheetImpl)sheet).nextConditionalId(); //ZSS-1251
	}
	
	@Override
	public SSheet getSheet() {
		return _sheet;
	}

	/* (non-Javadoc)
	 * @see org.zkoss.zss.model.SConditionalFormatting#getRegions()
	 */
	@Override
	public Set<CellRegion> getRegions() {
		return _regions;
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

	//ZSS-1251
	private Ref getRef(){
		return new ConditionalRefImpl(this);
	}
	@Override
	public void addRegion(CellRegion region) {
		Validations.argNotNull(region);
		if (this._regions == null) {
			this._regions = new LinkedHashSet<CellRegion>(2);
		}
		for (CellRegion regn : this._regions) {
			if (regn.contains(region)) return; // already in this DataValidation, let go
		}
		
		this._regions.add(region);
		
		// Add new ObjectRef into DependencyTable so we can extend/shrink/move
		Ref dependent = getRef();
		SBook book = _sheet.getBook();
		final DependencyTable dt = 
				((AbstractBookSeriesAdv) book.getBookSeries()).getDependencyTable();
		// prepare a dummy CellRef to enforce DataValidation reference dependency
		final Ref rgnRef = newDummyRef(_sheet.getSheetName(), region);
		dt.add(dependent, rgnRef);
		dt.add(rgnRef, dependent);
		
		ModelUpdateUtil.addRefUpdate(dependent);
	}
	
	//ZSS-1251
	@Override
	public void removeRegion(CellRegion region) {
		Validations.argNotNull(region);
		if (this._regions == null || this._regions.isEmpty()) return;
		
		List<CellRegion> newRegions = new ArrayList<CellRegion>();
		List<CellRegion> delRegions = new ArrayList<CellRegion>();
		final CellRegion leftop = this._regions.iterator().next();
		
		for (CellRegion regn : this._regions) {
			if (!regn.overlaps(region)) continue;
			List<CellRegion> diff = regn.diff(region);			
			newRegions.addAll(diff);
			delRegions.add(regn);
		}
		
		// no overlapping at all
		if (newRegions.isEmpty() && delRegions.isEmpty()) {
			return;
		}

		// remove old regions
		for (CellRegion regn : delRegions) {
			this._regions.remove(regn);
		}
		// add new split regions
		for (CellRegion regn : newRegions) {
			this._regions.add(regn);
		}
		// clear if empty
		if (this._regions.isEmpty()) {
			this._regions = null;
		}
		
		CellRegion target = null;
		if (this._regions != null) {
			//locate the new left-top region
			for (CellRegion diff0: this._regions) {
				if (target == null) {
					target = diff0;
				} else if (diff0.row < target.row) {
					target = diff0;
				} else if (diff0.row == target.row && diff0.column < target.column) {
					target = diff0;
				}
			}
		}
		
		// after the new region is set; the formula inside rule might have to be shifted
		int rowOff = target == null ? 0 : target.row - leftop.row;
		int colOff = target == null ? 0 : target.column - leftop.column;
		
		final String sheetName = _sheet.getSheetName();
		// manage the dependency table on rule
		Ref dependent = getRef();
		SBook book = _sheet.getBook();
		final DependencyTable dt = 
				((AbstractBookSeriesAdv) book.getBookSeries()).getDependencyTable();
		dt.clearDependents(dependent);
		for (CellRegion regn : delRegions) {
			dt.del(newDummyRef(sheetName, regn), dependent);
		}
		
		for (SConditionalFormattingRule rule : getRules()) {
			//ZSS-1251: shift rule Formula if any; 
			rule.shiftFormulas(rowOff, colOff);
			
			//ZSS-1251: clear the cache and force rebuild dependency table per rule
			rule.clearFormulaResultCache();			
		}
		
		ModelUpdateUtil.addRefUpdate(dependent);
	}
	
	// ZSS-1251
	private Ref newDummyRef(String sheetName, CellRegion regn) {
		return new RefImpl(_sheet.getBook().getBookName(), sheetName, 
				regn.row, regn.column, regn.lastRow, regn.lastColumn);
	}
	
	//ZSS-1251
	@Override
	public void renameSheet(String oldName, String newName) { //ZSS-648
		Validations.argNotNull(oldName);
		Validations.argNotNull(newName);
		if (oldName.equals(newName)) return; // nothing change, let go
		
		for (SConditionalFormattingRule rule : getRules()) {
			// remove old ObjectRef
			Ref dependent = ((ConditionalFormattingRuleImpl)rule).getRef(oldName);
			SBook book = _sheet.getBook();
			final DependencyTable dt = 
					((AbstractBookSeriesAdv) book.getBookSeries()).getDependencyTable();
			final Set<Ref> precedents = ((DependencyTableAdv)dt).getDirectPrecedents(dependent);
			if (precedents != null && this._regions != null) {
				for (CellRegion regn : this._regions) {
					precedents.remove(newDummyRef(oldName, regn));
				}
			}
			dt.clearDependents(dependent);
			
			// Add new ObjectRef into DependencyTable so we can extend/shrink/move
			dependent = ((ConditionalFormattingRuleImpl)rule).getRef(newName);  // new dependent because region have changed
			
			// ZSS-648
			// prepare new dummy CellRef to enforce ConditionalFormatting reference dependency
			if (this._regions != null) {
				for (CellRegion regn : this._regions) {
					dt.add(dependent, newDummyRef(newName, regn));
				}
			}
			
			// restore dependent precedents relation
			if (precedents != null) {
				for (Ref precedent: precedents) {
					dt.add(dependent, precedent);
				}
			}
		}
	}
	
	@Override
	public void destroy() {
		for (SConditionalFormattingRule rule: getRules()) {
			rule.destroy();
		}
		_sheet = null;
	}
	
	//ZSS-1251
	@Override
	public void copyFrom(SConditionalFormatting src, int rowOff, int colOff) {
		Validations.argInstance(src, ConditionalFormattingImpl.class);
		ConditionalFormattingImpl srcImpl = (ConditionalFormattingImpl)src;
		
		for (SConditionalFormattingRule rule : src.getRules()) {
			this.addConditionalFormattingRule(rule, rowOff, colOff);
		}
	}

	//ZSS-1251
	@Override
	public void setRegions(Set<CellRegion> regions) {
		_regions = new LinkedHashSet<CellRegion>(regions.size() * 4 / 3 + 1);
		for (CellRegion rgn : regions) {
			addRegion(rgn);
		}
	}

	//ZSS-1251
	public SConditionalFormattingRule addConditionalFormattingRule(SConditionalFormattingRule src, int rowOff, int colOff) {
		Validations.argInstance(src, ConditionalFormattingRuleImpl.class);
		ConditionalFormattingRuleImpl dstrule = new ConditionalFormattingRuleImpl(this);
		addRule(dstrule);
		if (src != null) {
			dstrule.copyFrom((ConditionalFormattingRuleImpl) src, rowOff, colOff);
		}
		return dstrule;
	}
	
	//ZSS-1251
	@Override
	public int getId() {
		return _id;
	}
	
	//ZSS-1251
	@Override
	public void clearFormulaResultCache() {
		for (SConditionalFormattingRule rule : getRules()) {
			rule.clearFormulaResultCache();
		}
	}
}
