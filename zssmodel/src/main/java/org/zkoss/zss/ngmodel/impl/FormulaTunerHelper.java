package org.zkoss.zss.ngmodel.impl;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.NChart;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.SheetRegion;
import org.zkoss.zss.ngmodel.chart.NChartData;
import org.zkoss.zss.ngmodel.chart.NGeneralChartData;
import org.zkoss.zss.ngmodel.chart.NSeries;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef.ObjectType;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;

/*package*/ class FormulaTunerHelper {
	final NBookSeries bookSeries;

	public FormulaTunerHelper(NBookSeries bookSeries) {
		this.bookSeries = bookSeries;
	}

	public void move(SheetRegion sheetRegion,Set<Ref> dependents,int rowOffset, int columnOffset) {
		//because of the chart shifting is for all chart, but the input dependent is on series,
		//so we need to collect the dependent for only shift chart once
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();
		
		for (Ref dependent : dependents) {
			if (dependent.getType() == RefType.CELL) {
				System.out.println(">>>Move Sheet Cell Formula: "+dependent);
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
		NBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		NChart chart =  sheet.getChart(dependent.getObjectIdPath()[0]);
		if(chart==null) return;
		NChartData d = chart.getData();
		if(!(d instanceof NGeneralChartData)) return;
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter;
		NGeneralChartData data = (NGeneralChartData)d;

		String catExpr = data.getCategoriesFormula();
		if(catExpr!=null){
			exprAfter = engine.move(catExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprAfter.hasError() && !catExpr.equals(exprAfter.getFormulaString())){
				System.out.println(">>>cat "+catExpr+" to "+exprAfter.hasError());
				data.setCategoriesFormula(exprAfter.getFormulaString());
			}
		}
		
		for(int i=0;i<data.getNumOfSeries();i++){
			NSeries series = data.getSeries(i);
			System.out.println(">>>move series "+series.getId());
			String nameExpr = series.getNameFormula();
			String xvalExpr = series.getXValuesFormula();
			String yvalExpr = series.getYValuesFormula();
			String zvalExpr = series.getZValuesFormula();
			
			boolean dirty = false;
			
			if(nameExpr!=null){
				exprAfter = engine.move(nameExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>nameExpr "+nameExpr+" to "+exprAfter.getFormulaString());
					dirty |= !nameExpr.equals(exprAfter.getFormulaString()); 
					nameExpr = exprAfter.getFormulaString();
					
				}
			}
			if(xvalExpr!=null){
				exprAfter = engine.move(xvalExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>xvalExpr "+xvalExpr+" to "+exprAfter.getFormulaString());
					dirty |= !xvalExpr.equals(exprAfter.getFormulaString());
					xvalExpr = exprAfter.getFormulaString();
					
				}
			}
			if(yvalExpr!=null){
				exprAfter = engine.move(yvalExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>yvalExpr "+yvalExpr+" to "+exprAfter.getFormulaString());
					dirty |= !yvalExpr.equals(exprAfter.getFormulaString());
					yvalExpr = exprAfter.getFormulaString();
					
				}
			}
			if(zvalExpr!=null){
				exprAfter = engine.move(zvalExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>zvalExpr "+zvalExpr+" to "+exprAfter.getFormulaString());
					dirty |= !zvalExpr.equals(exprAfter.getFormulaString());
					zvalExpr = exprAfter.getFormulaString();
				}
			}
			if(dirty){
				series.setXYZFormula(nameExpr, xvalExpr, yvalExpr, zvalExpr);
			}
			System.out.println(">>>end move series "+series.getId());
		}
		
		ModelUpdateUtil.addRefUpdate(dependent);
		
	}
	private void moveDataValidationRef(SheetRegion sheetRegion,ObjectRef dependent,int rowOffset, int columnOffset) {
		//TODO zss 3.5
		throw new UnsupportedOperationException("not implement yet");
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
		NBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		NCell cell = sheet.getCell(dependent.getRow(),
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
			System.out.println(">>>Extend Sheet Formula: "+dependent);
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
		throw new UnsupportedOperationException("not implement yet");
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
		throw new UnsupportedOperationException("not implement yet");
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
		NBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		NCell cell = sheet.getCell(dependent.getRow(),
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
		System.out.println(">>>>"+expr+" extend to "+exprAfter.getFormulaString());
	}	

	public void shrink(SheetRegion sheetRegion,Set<Ref> dependents, boolean horizontal) {
		//because of the chart shifting is for all chart, but the input dependent is on series,
		//so we need to collect the dependent for only shift chart once
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();		
		for (Ref dependent : dependents) {
			System.out.println(">>>Shrink Sheet Formula: "+dependent);
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
		throw new UnsupportedOperationException("not implement yet");
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
		throw new UnsupportedOperationException("not implement yet");
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
		NBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		NCell cell = sheet.getCell(dependent.getRow(),
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
		System.out.println(">>>>"+expr+" shrink to "+exprAfter.getFormulaString());
	}

	public void renameSheet(NBook book, String oldName, String newName,
			Set<Ref> dependents) {
		//because of the chart shifting is for all chart, but the input dependent is on series,
		//so we need to collect the dependent for only shift chart once
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();		
		for (Ref dependent : dependents) {
			System.out.println(">>>Rename Sheet Formula: "+dependent);
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
	
	private void renameSheetChartRef(NBook bookOfSheet, String oldName, String newName,ObjectRef dependent) {
		NBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null){//the sheet was renamed., get form newname if possible
			if(oldName.equals(dependent.getSheetName())){
				sheet = book.getSheetByName(newName);
			}
		}
		if(sheet==null) return;
		NChart chart =  sheet.getChart(dependent.getObjectIdPath()[0]);
		if(chart==null) return;
		NChartData d = chart.getData();
		if(!(d instanceof NGeneralChartData)) return;
		
		FormulaEngine engine = getFormulaEngine();
		FormulaExpression exprAfter;
		NGeneralChartData data = (NGeneralChartData)d;
		/*
		 * for sheet rename case, we should always update formula to make new dependency, shouln't ignore if the formula string is the same
		 * Note, in other move cell case, we could ignore to set same formula string
		 */
		
		String catExpr = data.getCategoriesFormula();
		if(catExpr!=null){
			exprAfter = engine.renameSheet(catExpr, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
			if(!exprAfter.hasError()){
				System.out.println(">>>cat "+catExpr+" to "+exprAfter.getFormulaString());
				data.setCategoriesFormula(exprAfter.getFormulaString());
			}
		}
		
		for(int i=0;i<data.getNumOfSeries();i++){
			NSeries series = data.getSeries(i);
			System.out.println(">rename series "+series.getId());
			String nameExpr = series.getNameFormula();
			String xvalExpr = series.getXValuesFormula();
			String yvalExpr = series.getYValuesFormula();
			String zvalExpr = series.getZValuesFormula();
			if(nameExpr!=null){
				exprAfter = engine.renameSheet(nameExpr, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>nameExpr "+nameExpr+" to "+exprAfter.getFormulaString());
					nameExpr = exprAfter.getFormulaString();
				}
			}
			if(xvalExpr!=null){
				exprAfter = engine.renameSheet(xvalExpr, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>xvalExpr "+xvalExpr+" to "+exprAfter.getFormulaString());
					xvalExpr = exprAfter.getFormulaString();
				}
			}
			if(yvalExpr!=null){
				exprAfter = engine.renameSheet(yvalExpr, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>yvalExpr "+yvalExpr+" to "+exprAfter.getFormulaString());
					yvalExpr = exprAfter.getFormulaString();
				}
			}
			if(zvalExpr!=null){
				exprAfter = engine.renameSheet(zvalExpr, bookOfSheet, oldName, newName,new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>zvalExpr "+zvalExpr+" to "+exprAfter.getFormulaString());
					zvalExpr = exprAfter.getFormulaString();
				}
			}
			series.setXYZFormula(nameExpr, xvalExpr, yvalExpr, zvalExpr);
			System.out.println(">end reanme series "+series.getId());
		}
		
		ModelUpdateUtil.addRefUpdate(dependent);
		
	}	
	
	private void renameSheetDataValidationRef(NBook bookOfSheet, String oldName, String newName,ObjectRef dependent) {
		//TODO zss 3.5
		throw new UnsupportedOperationException("not implement yet");
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

	private void renameSheetCellRef(NBook bookOfSheet, String oldName, String newName,Ref dependent) {
		NBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null){//the sheet was renamed., get form newname if possible
			if(oldName.equals(dependent.getSheetName())){
				sheet = book.getSheetByName(newName);
			}
		}
		if(sheet==null) return;
		NCell cell = sheet.getCell(dependent.getRow(),
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
		
		System.out.println(">>>>"+expr+" rename to "+exprAfter.getFormulaString());
	}	
}
