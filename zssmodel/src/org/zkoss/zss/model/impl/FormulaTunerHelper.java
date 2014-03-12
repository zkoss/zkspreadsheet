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

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

import org.zkoss.util.logging.Log;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeries;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SChart;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.chart.SChartData;
import org.zkoss.zss.model.chart.SGeneralChartData;
import org.zkoss.zss.model.chart.SSeries;
import org.zkoss.zss.model.sys.EngineFactory;
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
		
		for (Ref dependent : dependents) {
			if (dependent.getType() == RefType.CELL) {
				moveCellRef(sheetRegion,dependent,rowOffset,columnOffset);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					chartDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					validationDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}
			} else {// TODO another

			}
		}
		for (Ref dependent : chartDependents.values()) {
			moveChartRef(sheetRegion,(ObjectRef)dependent,rowOffset,columnOffset);
		}
		for (Ref dependent : validationDependents.values()) {
			moveDataValidationRef(sheetRegion,(ObjectRef)dependent,rowOffset,columnOffset);
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
			}
		}
		
		ModelUpdateUtil.addRefUpdate(dependent);
		
	}
	private void moveDataValidationRef(SheetRegion sheetRegion,ObjectRef dependent,int rowOffset, int columnOffset) {
		//TODO zss 3.5
//		NBook book = bookSeries.getBook(dependent.getBookName());
//		if(book==null) return;
//		NSheet sheet = book.getSheetByName(dependent.getSheetName());
//		if(sheet==null) return;
//		String[] ids = dependent.getObjectIdPath();
//		NDataValidation validation = sheet.getDataValidation(ids[0]);
//		if(validation!=null){
//			validation.clearFormulaResultCache();
//		}
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
		
		if(!expr.equals(exprAfter)){
			cell.setFormulaValue(exprAfter.getFormulaString());
			//don't need to notify cell change, cell will do
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
				
		for (Ref dependent : dependents) {
			if (dependent.getType() == RefType.CELL) {
				extendCellRef(sheetRegion,dependent,horizontal);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					chartDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					validationDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}
			} else {// TODO another

			}
		}
		for (Ref dependent : chartDependents.values()) {
			extendChartRef(sheetRegion,(ObjectRef)dependent,horizontal);
		}
		for (Ref dependent : validationDependents.values()) {
			extendDataValidationRef(sheetRegion,(ObjectRef)dependent,horizontal);
		}		
	}
	private void extendChartRef(SheetRegion sheetRegion,ObjectRef dependent, boolean horizontal) {
		//TODO zss 3.5
//		NBook book = bookSeries.getBook(dependent.getBookName());
//		if(book==null) return;
//		NSheet sheet = book.getSheetByName(dependent.getSheetName());
//		if(sheet==null) return;
//		String[] ids = dependent.getObjectIdPath();
//		NChart chart = sheet.getChart(ids[0]);
//		if(chart!=null){
//			chart.getData().clearFormulaResultCache();
//		}
	}
	private void extendDataValidationRef(SheetRegion sheetRegion,ObjectRef dependent, boolean horizontal) {
		//TODO zss 3.5
//		NBook book = bookSeries.getBook(dependent.getBookName());
//		if(book==null) return;
//		NSheet sheet = book.getSheetByName(dependent.getSheetName());
//		if(sheet==null) return;
//		String[] ids = dependent.getObjectIdPath();
//		NDataValidation validation = sheet.getDataValidation(ids[0]);
//		if(validation!=null){
//			validation.clearFormulaResultCache();
//		}
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
		if(!expr.equals(exprAfter)){
			cell.setFormulaValue(exprAfter.getFormulaString());
			//don't need to notify cell change, cell will do
		}
	}	

	public void shrink(SheetRegion sheetRegion,Set<Ref> dependents, boolean horizontal) {
		//because of the chart shifting is for all chart, but the input dependent is on series,
		//so we need to collect the dependent for only shift chart once
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();		
		for (Ref dependent : dependents) {
			if (dependent.getType() == RefType.CELL) {
				shrinkCellRef(sheetRegion,dependent,horizontal);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					chartDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					validationDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}
			} else {// TODO another

			}
		}
		for (Ref dependent : chartDependents.values()) {
			shrinkChartRef(sheetRegion,(ObjectRef)dependent,horizontal);
		}
		for (Ref dependent : validationDependents.values()) {
			shrinkDataValidationRef(sheetRegion,(ObjectRef)dependent,horizontal);
		}		
	}
	private void shrinkChartRef(SheetRegion sheetRegion,ObjectRef dependent, boolean horizontal) {
		//TODO zss 3.5
//		NBook book = bookSeries.getBook(dependent.getBookName());
//		if(book==null) return;
//		NSheet sheet = book.getSheetByName(dependent.getSheetName());
//		if(sheet==null) return;
//		String[] ids = dependent.getObjectIdPath();
//		NChart chart = sheet.getChart(ids[0]);
//		if(chart!=null){
//			chart.getData().clearFormulaResultCache();
//		}
	}
	private void shrinkDataValidationRef(SheetRegion sheetRegion,ObjectRef dependent, boolean horizontal) {
		//TODO zss 3.5
//		NBook book = bookSeries.getBook(dependent.getBookName());
//		if(book==null) return;
//		NSheet sheet = book.getSheetByName(dependent.getSheetName());
//		if(sheet==null) return;
//		String[] ids = dependent.getObjectIdPath();
//		NDataValidation validation = sheet.getDataValidation(ids[0]);
//		if(validation!=null){
//			validation.clearFormulaResultCache();
//		}
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
		if(!expr.equals(exprAfter)){
			cell.setFormulaValue(exprAfter.getFormulaString());
			//don't need to notify cell change, cell will do
		}
	}

	public void renameSheet(SBook book, String oldName, String newName,
			Set<Ref> dependents) {
		//because of the chart shifting is for all chart, but the input dependent is on series,
		//so we need to collect the dependent for only shift chart once
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();		
		for (Ref dependent : dependents) {
			if (dependent.getType() == RefType.CELL) {
				renameSheetCellRef(book,oldName,newName,dependent);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					chartDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					validationDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}
			} else {// TODO another

			}
		}
		for (Ref dependent : chartDependents.values()) {
			renameSheetChartRef(book,oldName,newName,(ObjectRef)dependent);
		}
		for (Ref dependent : validationDependents.values()) {
			renameSheetDataValidationRef(book,oldName,newName,(ObjectRef)dependent);
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
		//TODO zss 3.5
//		NBook book = bookSeries.getBook(dependent.getBookName());
//		if(book==null) return;
//		NSheet sheet = book.getSheetByName(dependent.getSheetName());
//		if(sheet==null) return;
//		String[] ids = dependent.getObjectIdPath();
//		NDataValidation validation = sheet.getDataValidation(ids[0]);
//		if(validation!=null){
//			validation.clearFormulaResultCache();
//		}
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
}
