/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.impl;

import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeries;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SChart;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SName;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.chart.SChartData;
import org.zkoss.zss.model.chart.SGeneralChartData;
import org.zkoss.zss.model.chart.SSeries;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.NameRef;
import org.zkoss.zss.model.sys.dependency.ObjectRef;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.dependency.ObjectRef.ObjectType;
import org.zkoss.zss.model.sys.dependency.Ref.RefType;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.FormulaParseContext;
/**
 * 
 * @author Dennis
 * @since 3.5.0
 */
/*package*/ class FormulaTunerHelper {
	private final SBookSeries _bookSeries;

	public FormulaTunerHelper(SBookSeries bookSeries) {
		this._bookSeries = bookSeries;
	}

	public void move(SheetRegion sheetRegion,Set<Ref> dependents,int rowOffset, int columnOffset) {
		//because of the chart shifting is for all chart, but the input dependent is on series,
		//so we need to collect the dependent for only shift chart once
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();
		Set<Ref> cellDependents = new LinkedHashSet<Ref>(); //ZSS-649
		Map<String,Ref> nameDependents = new LinkedHashMap<String, Ref>(); //ZSS-649
		Set<Ref> filterDependents = new LinkedHashSet<Ref>(); //ZSS-555
		
		splitDependents(dependents, cellDependents, chartDependents, validationDependents, nameDependents, filterDependents);
		
		for (Ref dependent : cellDependents) {
			moveCellRef(sheetRegion,dependent,rowOffset,columnOffset);
		}
		for (Ref dependent : chartDependents.values()) {
			moveChartRef(sheetRegion,(ObjectRef)dependent,rowOffset,columnOffset);
		}
		for (Ref dependent : validationDependents.values()) {
			moveDataValidationRef(sheetRegion,(ObjectRef)dependent,rowOffset,columnOffset);
		}
		for (Ref dependent : nameDependents.values()) {
			moveNameRef(sheetRegion,(NameRef)dependent,rowOffset,columnOffset);
		}
		//ZSS-555
		for (Ref dependent : filterDependents) {
			moveFilterRef(sheetRegion,(ObjectRef)dependent,rowOffset,columnOffset);
		}
	}
	private void moveChartRef(SheetRegion sheetRegion,ObjectRef dependent,int rowOffset, int columnOffset) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		SChart chart =  sheet.getChart(dependent.getObjectIdPath()[0]);
		if(chart==null) return;
		SChartData d = chart.getData();
		if(!(d instanceof SGeneralChartData)) return;
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter;
		SGeneralChartData data = (SGeneralChartData)d;

		String catExpr = data.getCategoriesFormula();
		if(catExpr!=null){
			exprAfter = engine.move(catExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprAfter.hasError() && !catExpr.equals(exprAfter.getFormulaString())){
				data.setCategoriesFormula(exprAfter.getFormulaString());
			}else{
				//zss-626, has to clear cache and notify ref update
				data.clearFormulaResultCache();
			}
		}
		
		for(int i=0;i<data.getNumOfSeries();i++){
			SSeries series = data.getSeries(i);
			String nameExpr = series.getNameFormula();
			String xvalExpr = series.getXValuesFormula();
			String yvalExpr = series.getYValuesFormula();
			String zvalExpr = series.getZValuesFormula();
			
			boolean dirty = false;
			
			if(nameExpr!=null){
				exprAfter = engine.move(nameExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					dirty |= !nameExpr.equals(exprAfter.getFormulaString()); 
					nameExpr = exprAfter.getFormulaString();
					
				}
			}
			if(xvalExpr!=null){
				exprAfter = engine.move(xvalExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					dirty |= !xvalExpr.equals(exprAfter.getFormulaString());
					xvalExpr = exprAfter.getFormulaString();
					
				}
			}
			if(yvalExpr!=null){
				exprAfter = engine.move(yvalExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					dirty |= !yvalExpr.equals(exprAfter.getFormulaString());
					yvalExpr = exprAfter.getFormulaString();
					
				}
			}
			if(zvalExpr!=null){
				exprAfter = engine.move(zvalExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					dirty |= !zvalExpr.equals(exprAfter.getFormulaString());
					zvalExpr = exprAfter.getFormulaString();
				}
			}
			if(dirty){
				series.setXYZFormula(nameExpr, xvalExpr, yvalExpr, zvalExpr);
			}else{
				//zss-626, has to clear cache and notify ref update
				series.clearFormulaResultCache();
			}
		}
		
		ModelUpdateUtil.addRefUpdate(dependent);
		
	}
	//ZSS-555
	private void moveFilterRef(SheetRegion sheetRegion,ObjectRef dependent,int rowOffset, int columnOffset) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet == null) {
			return;
		}
		SAutoFilter filter = sheet.getAutoFilter();
		if(filter == null) {
			return;
		}
		FormulaEngine engine = getFormulaEngine();
		
		// update AutoFilter's region
		final CellRegion region = filter.getRegion();
		final String area = region.getReferenceString();
			
		FormulaExpression expr2 = engine.move(area, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null)); //null ref, no trace dependence here
		if(!expr2.hasError() && !area.equals(expr2.getFormulaString())) {
			sheet.deleteAutoFilter();
			if ("#REF!".equals(expr2.getFormulaString())) { // should delete the region
				//delete
			} else { //delete then add
				final CellRegion region2 = new CellRegion(expr2.getFormulaString());
				sheet.createAutoFilter(region2);
			}
		}

		// notify filter change
		ModelUpdateUtil.addRefUpdate(dependent);
	}
	
	//ZSS-648
	private void moveDataValidationRef(SheetRegion sheetRegion,ObjectRef dependent,int rowOffset, int columnOffset) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet == null) {
			return;
		}
		SDataValidation validation = sheet.getDataValidation(dependent.getObjectIdPath()[0]);
		if(validation == null) {
			return;
		}
		FormulaEngine engine = getFormulaEngine();
		
		// update Validation's formula if any
		String f1 = validation.getFormula1();
		String f2 = validation.getFormula2();
		boolean changed = false;
		if (f1 != null) {
			FormulaExpression exprf1 = engine.move(f1, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprf1.hasError() && !f1.equals(exprf1.getFormulaString())) {
				f1 = exprf1.getFormulaString();
				changed = true;
			}
		}
		if (f2 != null) {
			FormulaExpression exprf2 = engine.move(f2, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprf2.hasError() && !f2.equals(exprf2.getFormulaString())) {
				f2 = exprf2.getFormulaString();
				changed = true;
			}
		}
		if (changed) {
			((AbstractDataValidationAdv)validation).setFormulas(f1, f2);
		} else {
			validation.clearFormulaResultCache();
		}
		
		// update Validation's region (sqref)
		for (CellRegion region : validation.getRegions()) {
			String sqref = region.getReferenceString();
			
			FormulaExpression expr2 = engine.move(sqref, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null)); //null ref, no trace dependence here
			if(!expr2.hasError() && !sqref.equals(expr2.getFormulaString())) {
				if ("#REF!".equals(expr2.getFormulaString())) { // should delete the region
					((AbstractDataValidationAdv)validation).removeRegion(region);
					if (validation.getRegions() == null) {
						sheet.deleteDataValidation(validation);
					}
				} else {
					region = new CellRegion(expr2.getFormulaString());
					((AbstractDataValidationAdv)validation).addRegion(region);
				}
			}
		}

		// notify chart change
		ModelUpdateUtil.addRefUpdate(dependent);
	}

	private void moveCellRef(SheetRegion sheetRegion,Ref dependent,int rowOffset, int columnOffset) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		SCell cell = sheet.getCell(dependent.getRow(),
				dependent.getColumn());
		if(cell.getType()!=CellType.FORMULA)
			return;//impossible
		
		String expr = cell.getFormulaValue();
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter = engine.move(expr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(cell, null));//null ref, no trace dependence here
		
		if(!expr.equals(exprAfter.getFormulaString())){
			cell.setFormulaValue(exprAfter.getFormulaString());
			//don't need to notify cell change, cell will do
		}else{
			//zss-626, has to clear cache and notify ref update
			cell.clearFormulaResultCache();
			ModelUpdateUtil.addRefUpdate(dependent);
		}
	}
	//ZSS-649
	private void moveNameRef(SheetRegion sheetRegion,NameRef dependent,int rowOffset, int columnOffset) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SName name = book.getNameByName(dependent.getNameName());
		if(name==null) return;
		SSheet sheet = book.getSheetByName(name.getRefersToSheetName());
		if(sheet==null) return;
		String expr = name.getRefersToFormula();
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter = engine.move(expr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
		
		if(!expr.equals(exprAfter.getFormulaString())){
			name.setRefersToFormula(exprAfter.getFormulaString());
			//don't need to notify name change, name will do
		}else{
			//zss-687, clear dependents's cache of NameRef
			clearFormulaCache(dependent);
			
			//zss-626, has to clear cache and notify ref update
			ModelUpdateUtil.addRefUpdate(dependent);
		}
	}

	private FormulaEngine engine;
	private FormulaEngine getFormulaEngine() {
		if(engine==null){
			engine = EngineFactory.getInstance().createFormulaEngine();
		}
		return engine;
	}
	
	public void extend(SheetRegion sheetRegion,Set<Ref> dependents, boolean horizontal) {
		//because of the chart shifting is for all chart, but the input dependent is on series,
		//so we need to collect the dependent for only shift chart once
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();
		Set<Ref> cellDependents = new LinkedHashSet<Ref>(); //ZSS-649
		Map<String,Ref> nameDependents = new LinkedHashMap<String, Ref>(); //ZSS-649
		Set<Ref> filterDependents = new LinkedHashSet<Ref>(); //ZSS-555
		
		splitDependents(dependents, cellDependents, chartDependents, validationDependents, nameDependents, filterDependents);
		
		for (Ref dependent : cellDependents) {
			extendCellRef(sheetRegion,dependent,horizontal);
		}
		for (Ref dependent : chartDependents.values()) {
			extendChartRef(sheetRegion,(ObjectRef)dependent,horizontal);
		}
		for (Ref dependent : validationDependents.values()) {
			extendDataValidationRef(sheetRegion,(ObjectRef)dependent,horizontal);
		}		
		for (Ref dependent : nameDependents.values()) {
			extendNameRef(sheetRegion,(NameRef)dependent,horizontal);
		}
		//ZSS-555
		for (Ref dependent : filterDependents) {
			extendFilterRef(sheetRegion,(ObjectRef)dependent,horizontal);
		}
		
	}

	private void extendChartRef(SheetRegion sheetRegion, ObjectRef dependent, boolean horizontal) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet == null) {
			return;
		}
		SChart chart = sheet.getChart(dependent.getObjectIdPath()[0]);
		if(chart == null) {
			return;
		}
		SChartData d = chart.getData();
		if(!(d instanceof SGeneralChartData)) {
			return;
		}
		SGeneralChartData data = (SGeneralChartData)d;
		FormulaEngine engine = getFormulaEngine();
		
		// extend series formula
		for(int i = 0; i < data.getNumOfSeries(); ++i) {
			SSeries series = data.getSeries(i);
			if(series != null) {
				boolean changed = false;
				series.clearFormulaResultCache();
				String nf = series.getNameFormula();
				String xf = series.getXValuesFormula();
				String yf = series.getYValuesFormula();
				String zf = series.getZValuesFormula();
				if(nf != null) {
					FormulaExpression expr2 = engine.extend(nf, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
					if(!expr2.hasError() && !nf.equals(expr2.getFormulaString())) {
						nf = expr2.getFormulaString();
						changed = true;
					}
				}
				if(xf != null) {
					FormulaExpression expr2 = engine.extend(xf, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
					if(!expr2.hasError() && !xf.equals(expr2.getFormulaString())) {
						xf = expr2.getFormulaString();
						changed = true;
					}
				}
				if(yf != null) {
					FormulaExpression expr2 = engine.extend(yf, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
					if(!expr2.hasError() && !yf.equals(expr2.getFormulaString())) {
						yf = expr2.getFormulaString();
						changed = true;
					}
				}
				if(zf != null) {
					FormulaExpression expr2 = engine.extend(zf, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
					if(!expr2.hasError() && !zf.equals(expr2.getFormulaString())) {
						zf = expr2.getFormulaString();
						changed = true;
					}
				}
				if(changed) {
					series.setXYZFormula(nf, xf, yf, zf);
				}else{
					//zss-626, has to clear cache and notify ref update
					series.clearFormulaResultCache();
				}
			}
		}
		
		// extend categories formula
		String expr = data.getCategoriesFormula();
		if(expr != null) {
			FormulaExpression exprAfter = engine.extend(expr, sheetRegion,horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprAfter.hasError() && !expr.equals(exprAfter.getFormulaString())) {
				data.setCategoriesFormula(exprAfter.getFormulaString());
			}else{
				//zss-626, has to clear cache and notify ref update
				data.clearFormulaResultCache();
			}
		}
		
		// notify chart change
		ModelUpdateUtil.addRefUpdate(dependent);
	}
	//ZSS-648
	private void extendDataValidationRef(SheetRegion sheetRegion,ObjectRef dependent, boolean horizontal) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet == null) {
			return;
		}
		SDataValidation validation = sheet.getDataValidation(dependent.getObjectIdPath()[0]);
		if(validation == null) {
			return;
		}
		FormulaEngine engine = getFormulaEngine();
		
		// update Validation's formula if any
		String f1 = validation.getFormula1();
		String f2 = validation.getFormula2();
		boolean changed = false;
		if (f1 != null) {
			FormulaExpression exprf1 = engine.extend(f1, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprf1.hasError() && !f1.equals(exprf1.getFormulaString())) {
				f1 = exprf1.getFormulaString();
				changed = true;
			}
		}
		if (f2 != null) {
			FormulaExpression exprf2 = engine.extend(f2, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprf2.hasError() && !f2.equals(exprf2.getFormulaString())) {
				f2 = exprf2.getFormulaString();
				changed = true;
			}
		}
		if (changed) {
			((AbstractDataValidationAdv)validation).setFormulas(f1, f2);
		} else {
			validation.clearFormulaResultCache();
		}
		
		// update Validation's region (sqref)
		for (CellRegion region : validation.getRegions()) {
			String sqref = region.getReferenceString();
			
			FormulaExpression expr2 = engine.extend(sqref, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!expr2.hasError() && !sqref.equals(expr2.getFormulaString())) {
				if ("#REF!".equals(expr2.getFormulaString())) { // should delete the region
					((AbstractDataValidationAdv)validation).removeRegion(region);
					if (validation.getRegions() == null) {
						sheet.deleteDataValidation(validation);
					}
				} else {
					region = new CellRegion(expr2.getFormulaString());
					((AbstractDataValidationAdv)validation).addRegion(region);
				}
			}
		}

		// notify validation change
		ModelUpdateUtil.addRefUpdate(dependent);
	}

	private void extendFilterRef(SheetRegion sheetRegion,ObjectRef dependent, boolean horizontal) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet == null) {
			return;
		}
		SAutoFilter filter = sheet.getAutoFilter();
		if(filter == null) {
			return;
		}
		FormulaEngine engine = getFormulaEngine();
		
		// update AutoFilter's region
		CellRegion region = filter.getRegion();
		String area = region.getReferenceString();
		
		FormulaExpression expr2 = engine.extend(area, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
		if(!expr2.hasError() && !area.equals(expr2.getFormulaString())) {
			sheet.deleteAutoFilter();
			if ("#REF!".equals(expr2.getFormulaString())) { // should delete the region
				// delete
			} else { // delete than add
				final CellRegion region2 = new CellRegion(expr2.getFormulaString());
				sheet.createAutoFilter(region2);
			}
		}

		// notify filter change
		ModelUpdateUtil.addRefUpdate(dependent);
	}
	
	private void extendCellRef(SheetRegion sheetRegion,Ref dependent, boolean horizontal) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		SCell cell = sheet.getCell(dependent.getRow(),
				dependent.getColumn());
		if(cell.getType()!=CellType.FORMULA)
			return;//impossible
		
		String expr = cell.getFormulaValue();
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter = engine.extend(expr, sheetRegion,horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
		if(!expr.equals(exprAfter.getFormulaString())){
			cell.setFormulaValue(exprAfter.getFormulaString());
			//don't need to notify cell change, cell will do
		}else{
			//zss-626, has to clear cache and notify ref update
			cell.clearFormulaResultCache();
			ModelUpdateUtil.addRefUpdate(dependent);
		}
	}

	//ZSS-649
	private void extendNameRef(SheetRegion sheetRegion, NameRef dependent, boolean horizontal) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SName name = book.getNameByName(dependent.getNameName());
		if (name == null) return;
		SSheet sheet = book.getSheetByName(name.getRefersToSheetName());
		if(sheet==null) return;
		String expr = name.getRefersToFormula();
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter = engine.extend(expr, sheetRegion,horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
		if(!expr.equals(exprAfter.getFormulaString())){
			name.setRefersToFormula(exprAfter.getFormulaString());
			//don't need to notify cell change, cell will do
		}else{
			//zss-687, clear dependents's cache of NameRef
			clearFormulaCache(dependent);
			
			//zss-626, has to clear cache and notify ref update
			ModelUpdateUtil.addRefUpdate(dependent);
		}
	}
	
	public void shrink(SheetRegion sheetRegion, Set<Ref> dependents, boolean horizontal) {
		//because of the chart shifting is for all chart, but the input dependent is on series,
		//so we need to collect the dependent for only shift chart once
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();
		Set<Ref> cellDependents = new LinkedHashSet<Ref>(); //ZSS-649
		Map<String,Ref> nameDependents = new LinkedHashMap<String, Ref>(); //ZSS-649
		Set<Ref> filterDependents = new LinkedHashSet<Ref>(); //ZSS-555
		
		splitDependents(dependents, cellDependents, chartDependents, validationDependents, nameDependents, filterDependents);
		
		for (Ref dependent : cellDependents) {
			shrinkCellRef(sheetRegion,dependent,horizontal);
		}
		for (Ref dependent : chartDependents.values()) {
			shrinkChartRef(sheetRegion,(ObjectRef)dependent,horizontal);
		}
		for (Ref dependent : validationDependents.values()) {
			shrinkDataValidationRef(sheetRegion,(ObjectRef)dependent,horizontal);
		}
		//ZSS-649
		for (Ref dependent : nameDependents.values()) {
			shrinkNameRef(sheetRegion, (NameRef) dependent, horizontal);
		}
		//ZSS-555
		for (Ref dependent : filterDependents) {
			shrinkFilterRef(sheetRegion,(ObjectRef)dependent,horizontal);
		}
	}
	private void shrinkChartRef(SheetRegion sheetRegion,ObjectRef dependent, boolean horizontal) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet == null) {
			return;
		}
		SChart chart = sheet.getChart(dependent.getObjectIdPath()[0]);
		if(chart == null) {
			return;
		}
		SChartData d = chart.getData();
		if(!(d instanceof SGeneralChartData)) {
			return;
		}
		SGeneralChartData data = (SGeneralChartData)d;
		FormulaEngine engine = getFormulaEngine();
		
		// shrink series formula
		for(int i = 0; i < data.getNumOfSeries(); ++i) {
			SSeries series = data.getSeries(i);
			if(series != null) {
				boolean changed = false;
				series.clearFormulaResultCache();
				String nf = series.getNameFormula();
				String xf = series.getXValuesFormula();
				String yf = series.getYValuesFormula();
				String zf = series.getZValuesFormula();
				if(nf != null) {
					FormulaExpression expr2 = engine.shrink(nf, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
					if(!expr2.hasError() && !nf.equals(expr2.getFormulaString())) {
						nf = expr2.getFormulaString();
						changed = true;
					}
				}
				if(xf != null) {
					FormulaExpression expr2 = engine.shrink(xf, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
					if(!expr2.hasError() && !xf.equals(expr2.getFormulaString())) {
						xf = expr2.getFormulaString();
						changed = true;
					}
				}
				if(yf != null) {
					FormulaExpression expr2 = engine.shrink(yf, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
					if(!expr2.hasError() && !yf.equals(expr2.getFormulaString())) {
						yf = expr2.getFormulaString();
						changed = true;
					}
				}
				if(zf != null) {
					FormulaExpression expr2 = engine.shrink(zf, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
					if(!expr2.hasError() && !zf.equals(expr2.getFormulaString())) {
						zf = expr2.getFormulaString();
						changed = true;
					}
				}
				if(changed) {
					series.setXYZFormula(nf, xf, yf, zf);
				}else{
					//zss-626, has to clear cache and notify ref update
					series.clearFormulaResultCache();
				}
			}
		}
		
		// shrink categories formula
		String expr = data.getCategoriesFormula();
		if(expr != null) {
			FormulaExpression exprAfter = engine.shrink(expr, sheetRegion,horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprAfter.hasError() && !expr.equals(exprAfter.getFormulaString())) {
				data.setCategoriesFormula(exprAfter.getFormulaString());
			}else{
				//zss-626, has to clear cache and notify ref update
				data.clearFormulaResultCache();
			}
		}
		
		// notify chart change
		ModelUpdateUtil.addRefUpdate(dependent);
	}
	//ZSS-648
	private void shrinkDataValidationRef(SheetRegion sheetRegion,ObjectRef dependent, boolean horizontal) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet == null) {
			return;
		}
		SDataValidation validation = sheet.getDataValidation(dependent.getObjectIdPath()[0]);
		if(validation == null) {
			return;
		}
		FormulaEngine engine = getFormulaEngine();
		
		// update Validation's formula if any
		String f1 = validation.getFormula1();
		String f2 = validation.getFormula2();
		boolean changed = false;
		if (f1 != null) {
			FormulaExpression exprf1 = engine.shrink(f1, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprf1.hasError() && !f1.equals(exprf1.getFormulaString())) {
				f1 = exprf1.getFormulaString();
				changed = true;
			}
		}
		if (f2 != null) {
			FormulaExpression exprf2 = engine.shrink(f2, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprf2.hasError() && !f2.equals(exprf2.getFormulaString())) {
				f2 = exprf2.getFormulaString();
				changed = true;
			}
		}
		if (changed) {
			((AbstractDataValidationAdv)validation).setFormulas(f1, f2);
		} else {
			validation.clearFormulaResultCache();
		}
		
		// update Validation's region (sqref)
		for (CellRegion region : validation.getRegions()) {
			String sqref = region.getReferenceString();
			
			FormulaExpression expr2 = engine.shrink(sqref, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!expr2.hasError() && !sqref.equals(expr2.getFormulaString())) {
				if ("#REF!".equals(expr2.getFormulaString())) { // should delete the region
					((AbstractDataValidationAdv)validation).removeRegion(region);
					if (validation.getRegions() == null) {
						sheet.deleteDataValidation(validation);
					}
				} else {
					region = new CellRegion(expr2.getFormulaString());
					((AbstractDataValidationAdv)validation).addRegion(region);
				}
			}
		}

		// notify chart change
		ModelUpdateUtil.addRefUpdate(dependent);
	}

	//ZSS-555
	private void shrinkFilterRef(SheetRegion sheetRegion,ObjectRef dependent, boolean horizontal) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet == null) {
			return;
		}
		SAutoFilter filter = sheet.getAutoFilter();
		if(filter == null) {
			return;
		}
		FormulaEngine engine = getFormulaEngine();
		
		// update AutoFilter's region
		CellRegion region = filter.getRegion();
		String area = region.getReferenceString();
		
		FormulaExpression expr2 = engine.shrink(area, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
		if(!expr2.hasError() && !area.equals(expr2.getFormulaString())) {
			sheet.deleteAutoFilter();
			if ("#REF!".equals(expr2.getFormulaString())) { // should delete the region
				//delete
			} else { // delete then add
				final CellRegion region2 = new CellRegion(expr2.getFormulaString());
				sheet.createAutoFilter(region2);
			}
		}

		// notify filter change
		ModelUpdateUtil.addRefUpdate(dependent);
	}

	private void shrinkCellRef(SheetRegion sheetRegion,Ref dependent, boolean horizontal) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		SCell cell = sheet.getCell(dependent.getRow(),
				dependent.getColumn());
		if(cell.getType()!=CellType.FORMULA)
			return;//impossible
		
		String expr = cell.getFormulaValue();
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter = engine.shrink(expr, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
		if(!expr.equals(exprAfter.getFormulaString())){
			cell.setFormulaValue(exprAfter.getFormulaString());
			//don't need to notify cell change, cell will do
		}else{
			//zss-626, has to clear cache and notify ref update
			cell.clearFormulaResultCache();
			ModelUpdateUtil.addRefUpdate(dependent);
		}
	}

	//ZSS-649
	private void shrinkNameRef(SheetRegion sheetRegion, NameRef dependent, boolean horizontal) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SName name = book.getNameByName(dependent.getNameName());
		if(name==null) return;
		SSheet sheet = book.getSheetByName(name.getRefersToSheetName());
		if(sheet==null) return;
		String expr = name.getRefersToFormula();
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter = engine.shrink(expr, sheetRegion, horizontal, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
		if(!expr.equals(exprAfter.getFormulaString())){
			name.setRefersToFormula(exprAfter.getFormulaString());
			//don't need to notify name change, setRefersToFormula will do
		}else{
			//zss-687, clear dependents's cache of NameRef
			clearFormulaCache(dependent);
			
			ModelUpdateUtil.addRefUpdate(dependent);
		}
	}
	
	public void renameSheet(SBook book, String oldName, String newName,
			Set<Ref> dependents) {
		//because of the chart shifting is for all chart, but the input dependent is on series,
		//so we need to collect the dependent for only shift chart once
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();
		Set<Ref> cellDependents = new LinkedHashSet<Ref>(); //ZSS-649
		Map<String,Ref> nameDependents = new LinkedHashMap<String, Ref>(); //ZSS-649
		Set<Ref> filterDependents = new LinkedHashSet<Ref>(); //ZSS-555
		
		splitDependents(dependents, cellDependents, chartDependents, validationDependents, nameDependents, filterDependents);
		
		for (Ref dependent : cellDependents) {
			renameSheetCellRef(book,oldName,newName,dependent);
		}
		for (Ref dependent : chartDependents.values()) {
			renameSheetChartRef(book,oldName,newName,(ObjectRef)dependent);
		}
		for (Ref dependent : validationDependents.values()) {
			renameSheetDataValidationRef(book,oldName,newName,(ObjectRef)dependent);
		}	
		//ZSS-649
		for (Ref dependent : nameDependents.values()) {
			renameSheetNameRef(book,oldName,newName,(NameRef)dependent);
		}
		//ZSS-555
		for (Ref dependent : filterDependents) {
			renameSheetFilterRef(book,oldName,newName,(ObjectRef)dependent);
		}
	}	
	
	private void renameSheetChartRef(SBook bookOfSheet, String oldName, String newName,ObjectRef dependent) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null){//the sheet was renamed., get form newname if possible
			if(oldName.equals(dependent.getSheetName())){
				sheet = book.getSheetByName(newName);
			}
		}
		if(sheet==null) return;
		SChart chart =  sheet.getChart(dependent.getObjectIdPath()[0]);
		if(chart==null) return;
		SChartData d = chart.getData();
		if(!(d instanceof SGeneralChartData)) return;
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter;
		SGeneralChartData data = (SGeneralChartData)d;
		/*
		 * for sheet rename case, we should always update formula to make new dependency, shouln't ignore if the formula string is the same
		 * Note, in other move cell case, we could ignore to set same formula string
		 */
		
		String catExpr = data.getCategoriesFormula();
		if(catExpr!=null){
			exprAfter = engine.renameSheet(catExpr, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprAfter.hasError()){
				data.setCategoriesFormula(exprAfter.getFormulaString());
			}
		}
		
		for(int i=0;i<data.getNumOfSeries();i++){
			SSeries series = data.getSeries(i);
			String nameExpr = series.getNameFormula();
			String xvalExpr = series.getXValuesFormula();
			String yvalExpr = series.getYValuesFormula();
			String zvalExpr = series.getZValuesFormula();
			if(nameExpr!=null){
				exprAfter = engine.renameSheet(nameExpr, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					nameExpr = exprAfter.getFormulaString();
				}
			}
			if(xvalExpr!=null){
				exprAfter = engine.renameSheet(xvalExpr, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					xvalExpr = exprAfter.getFormulaString();
				}
			}
			if(yvalExpr!=null){
				exprAfter = engine.renameSheet(yvalExpr, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					yvalExpr = exprAfter.getFormulaString();
				}
			}
			if(zvalExpr!=null){
				exprAfter = engine.renameSheet(zvalExpr, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					zvalExpr = exprAfter.getFormulaString();
				}
			}
			series.setXYZFormula(nameExpr, xvalExpr, yvalExpr, zvalExpr);
		}
		
		ModelUpdateUtil.addRefUpdate(dependent);
		
	}	
	
	private void renameSheetDataValidationRef(SBook bookOfSheet, String oldName, String newName,ObjectRef dependent) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null){//the sheet was renamed., get form newname if possible
			if(oldName.equals(dependent.getSheetName())){
				sheet = book.getSheetByName(newName);
			}
		}
		if(sheet == null) {
			return;
		}
		SDataValidation validation = sheet.getDataValidation(dependent.getObjectIdPath()[0]);
		if(validation == null) {
			return;
		}
		FormulaEngine engine = getFormulaEngine();
		
		// update Validation's formula if any
		String f1 = validation.getFormula1();
		String f2 = validation.getFormula2();
		boolean changed = false;
		if (f1 != null) {
			FormulaExpression exprf1 = engine.renameSheet(f1, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprf1.hasError() && !f1.equals(exprf1.getFormulaString())) {
				f1 = exprf1.getFormulaString();
				changed = true;
			}
		}
		if (f2 != null) {
			FormulaExpression exprf2 = engine.renameSheet(f2, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprf2.hasError() && !f2.equals(exprf2.getFormulaString())) {
				f2 = exprf2.getFormulaString();
				changed = true;
			}
		}
		if (changed) {
			((AbstractDataValidationAdv)validation).setFormulas(f1, f2);
		} else {
			validation.clearFormulaResultCache();
		}
		
		// update Validation's region (sqref)
		((AbstractDataValidationAdv)validation).renameSheet(oldName, newName);

		// notify chart change
		ModelUpdateUtil.addRefUpdate(dependent);
	}

	//ZSS-555
	private void renameSheetFilterRef(SBook bookOfSheet, String oldName, String newName,ObjectRef dependent) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null){//the sheet was renamed., get form newname if possible
			if(oldName.equals(dependent.getSheetName())){
				sheet = book.getSheetByName(newName);
			}
		}
		if(sheet == null) {
			return;
		}
		SAutoFilter filter = sheet.getAutoFilter();
		if(filter == null) {
			return;
		}
		
		// update AutoFilter's region
		((AbstractAutoFilterAdv)filter).renameSheet(book, oldName, newName);

		// notify filter change
		ModelUpdateUtil.addRefUpdate(dependent);
	}
	
	private void renameSheetCellRef(SBook bookOfSheet, String oldName, String newName,Ref dependent) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null){//the sheet was renamed., get form newname if possible
			if(oldName.equals(dependent.getSheetName())){
				sheet = book.getSheetByName(newName);
			}
		}
		if(sheet==null) return;
		SCell cell = sheet.getCell(dependent.getRow(),
				dependent.getColumn());
		if(cell.getType()!=CellType.FORMULA)
			return;//impossible
		
		/*
		 * for sheet rename case, we should always update formula to make new dependency, shouln't ignore if the formula string is the same
		 * Note, in other move cell case, we could ignore to set same formula string
		 */
		
		String expr = cell.getFormulaValue();
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter = engine.renameSheet(expr, bookOfSheet, oldName, newName,new FormulaParseContext(cell, null));//null ref, no trace dependence here
		
		cell.setFormulaValue(exprAfter.getFormulaString());
		//don't need to notify cell change, cell will do
	}	

	//ZSS-649
	private void renameSheetNameRef(SBook bookOfSheet, String oldName, String newName, NameRef dependent) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SName name = book.getNameByName(dependent.getNameName());
		if(name==null) return;
		SSheet sheet = book.getSheetByName(name.getRefersToSheetName());
		if(sheet==null){//the sheet was renamed., get form newname if possible
			if(oldName.equals(name.getRefersToSheetName())){
				sheet = book.getSheetByName(newName);
			}
		}
		if(sheet==null) return;
		
		/*
		 * for sheet rename case, we should always update formula to make new dependency, shouln't ignore if the formula string is the same
		 * Note, in other move cell case, we could ignore to set same formula string
		 */
		
		String expr = name.getRefersToFormula();
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter = engine.renameSheet(expr, bookOfSheet, oldName, newName, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
		
		name.setRefersToFormula(exprAfter.getFormulaString());
		//don't need to notify cell change, cell will do
	}	
	
	// ZSS-661
	public void renameName(SBook book, String oldName, String newName,
			Set<Ref> dependents) {
		for (Ref dependent : dependents) {
			if (dependent.getType() == RefType.CELL) {
				renameNameCellRef(book, oldName, newName, dependent);
			}
		}
	}	

	private void renameNameCellRef(SBook bookOfSheet, String oldName, String newName, Ref dependent) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if (book == null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if (sheet == null) return;
		SCell cell = sheet.getCell(dependent.getRow(), dependent.getColumn());
		if(cell.getType() != CellType.FORMULA)
			return;//impossible
		
		/*
		 * for Name rename case, we should always update formula to make new 
		 * dependency, shouln't ignore if the formula string is the same
		 * Note, in other move cell case, we could ignore to set same formula string
		 */
		String expr = cell.getFormulaValue();
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter = 
				engine.renameName(expr, bookOfSheet, oldName, newName, new FormulaParseContext(cell, null));
		
		cell.setFormulaValue(exprAfter.getFormulaString());
		//don't need to notify cell change, cell will do
	}

	private void splitDependents(final Set<Ref> dependents,
			final Set<Ref> cellDependents,
			final Map<String, Ref> chartDependents,
			final Map<String, Ref> validationDependents,
			final Map<String, Ref> nameDependents,
			final Set<Ref> filterDependents) {
		
		for (Ref dependent : dependents) {
			RefType type = dependent.getType();
			if (type == RefType.CELL) {
				cellDependents.add(dependent);
			} else if (type == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					chartDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					validationDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.AUTO_FILTER){
					filterDependents.add(dependent);
				}
			} else if (type == RefType.NAME) { //ZSS-649
				nameDependents.put(((NameRef)dependent).toString(), dependent);
			} else {// TODO another

			}
		}
	}

	//ZSS-687
	private void clearFormulaCache(NameRef precedent) {
		Map<String,Ref> chartDependents  = new HashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new HashMap<String, Ref>();
		Set<Ref> nameDependents  = new HashSet<Ref>();
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv) _bookSeries;
		DependencyTable dt = bs.getDependencyTable();
		
		clearFormulaCache(precedent, dt, chartDependents, validationDependents, nameDependents);

		for (Ref dependent : chartDependents.values()) {
			clearFormulaCacheChartRef((ObjectRef)dependent);
		}
		for (Ref dependent : validationDependents.values()) {
			clearFormulaCacheDataValidationRef((ObjectRef)dependent);
		}		
	}

	// ZSS-687
	private void clearFormulaCache(NameRef precedent,
			DependencyTable dt, 
			Map<String,Ref> chartDependents,
			Map<String,Ref> validationDependents,
			Set<Ref> nameDependents) {
		
		Set<Ref> dependents = dt.getDependents(precedent);
		
		for (Ref dependent : dependents) {
			RefType type = dependent.getType(); 
			if (type == RefType.CELL) {
				clearFormulaCacheCellRef(dependent);
			} else if (type == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					chartDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					validationDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}
			} else if (type == RefType.NAME) {
				if (!nameDependents.contains(dependent)) {
					nameDependents.add(dependent);
					// recursive back when NameRef depends on NameRef
					clearFormulaCache((NameRef)dependent, dt, 
						chartDependents, validationDependents, nameDependents); 
				}
			} else {// TODO another

			}
		}
	}

	// ZSS-687, ZSS-648
	private void clearFormulaCacheDataValidationRef(ObjectRef dependent) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet == null) {
			return;
		}
		SDataValidation validation = sheet.getDataValidation(dependent.getObjectIdPath()[0]);
		if(validation == null) {
			return;
		}
		
		validation.clearFormulaResultCache();
		ModelUpdateUtil.addRefUpdate(dependent);
	}
	
	// ZSS-687
	private void clearFormulaCacheChartRef(ObjectRef dependent) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book == null) {
			return;
		}
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet == null) {
			return;
		}
		SChart chart = sheet.getChart(dependent.getObjectIdPath()[0]);
		if(chart == null) {
			return;
		}
		SChartData d = chart.getData();
		if(!(d instanceof SGeneralChartData)) {
			return;
		}
		SGeneralChartData data = (SGeneralChartData)d;

		data.clearFormulaResultCache();
		ModelUpdateUtil.addRefUpdate(dependent);
	}

	// ZSS-687
	private void clearFormulaCacheCellRef(Ref dependent) {
		SBook book = _bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		SCell cell = sheet.getCell(dependent.getRow(),
				dependent.getColumn());
		if(cell.getType()!=CellType.FORMULA)
			return;//impossible
		
		//zss-626, has to clear cache and notify ref update
		cell.clearFormulaResultCache();
		ModelUpdateUtil.addRefUpdate(dependent);
	}
}
