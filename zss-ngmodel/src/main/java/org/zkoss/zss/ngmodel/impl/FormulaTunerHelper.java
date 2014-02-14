package org.zkoss.zss.ngmodel.impl;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

import org.zkoss.zss.ngapi.impl.DependentUpdateCollector;
import org.zkoss.zss.ngmodel.ModelEvents;
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
import org.zkoss.zss.ngmodel.impl.chart.GeneralChartDataImpl;
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
	final SheetRegion sheetRegion;

	public FormulaTunerHelper(AbstractBookSeriesAdv bookSeries,
			SheetRegion sheetRegion) {
		this.bookSeries = bookSeries;
		this.sheetRegion = sheetRegion;
	}

	public void move(Set<Ref> dependents,int rowOffset, int columnOffset) {
		//TODO collect to same chart/validation and only for once
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>(); 
		for (Ref dependent : dependents) {
			if (dependent.getType() == RefType.CELL) {
				System.out.println(">>>Move Sheet Cell Formula: "+dependent);
				moveCellRef(dependent,rowOffset,columnOffset);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					chartDependents.put(((ObjectRef)dependent).getObjectIdPath()[0], dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					System.out.println(">>>Move Sheet Validation Formula: "+dependent);
					moveDataValidationRef((ObjectRef)dependent,rowOffset,columnOffset);
				}
			} else {// TODO another

			}
		}
		for (Ref dependent : chartDependents.values()) {
			moveChartRef((ObjectRef)dependent,rowOffset,columnOffset);
		}
	}
	private void moveChartRef(ObjectRef dependent,int rowOffset, int columnOffset) {
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
			
			if(nameExpr!=null){
				exprAfter = engine.move(nameExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>nameExpr "+nameExpr+" to "+exprAfter.getFormulaString());
					nameExpr = exprAfter.getFormulaString();
				}
			}
			if(xvalExpr!=null){
				exprAfter = engine.move(xvalExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>xvalExpr "+xvalExpr+" to "+exprAfter.getFormulaString());
					xvalExpr = exprAfter.getFormulaString();
				}
			}
			if(yvalExpr!=null){
				exprAfter = engine.move(yvalExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>yvalExpr "+yvalExpr+" to "+exprAfter.getFormulaString());
					yvalExpr = exprAfter.getFormulaString();
				}
			}
			if(zvalExpr!=null){
				exprAfter = engine.move(zvalExpr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
				if(!exprAfter.hasError()){
					System.out.println(">>>zvalExpr "+zvalExpr+" to "+exprAfter.getFormulaString());
					zvalExpr = exprAfter.getFormulaString();
				}
			}
			series.setXYZFormula(nameExpr, xvalExpr, yvalExpr, zvalExpr);
			System.out.println(">>>end move series "+series.getId());
		}
		
		ModelUpdateUtil.addDepednentUpdate(sheet, dependent);
		
	}
	private void moveDataValidationRef(ObjectRef dependent,int rowOffset, int columnOffset) {
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

	private void moveCellRef(Ref dependent,int rowOffset, int columnOffset) {
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
		FormulaExpression exprAfter = engine.move(expr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
		cell.setFormulaValue(exprAfter.getFormulaString());
		//don't need to notify cell change, cell will do
	}

	private FormulaEngine engine;
	private FormulaEngine getFormulaEngine() {
		if(engine==null){
			engine = EngineFactory.getInstance().createFormulaEngine();
		}
		return engine;
	}
	
	public void extend(Set<Ref> dependents, boolean horizontal) {
		for (Ref dependent : dependents) {
			System.out.println(">>>Extend Sheet Formula: "+dependent);
			if (dependent.getType() == RefType.CELL) {
				extendCellRef(dependent,horizontal);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					extendChartRef((ObjectRef)dependent,horizontal);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					extendDataValidationRef((ObjectRef)dependent,horizontal);
				}
			} else {// TODO another

			}
		}
	}
	private void extendChartRef(ObjectRef dependent, boolean horizontal) {
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
	private void extendDataValidationRef(ObjectRef dependent, boolean horizontal) {
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

	private void extendCellRef(Ref dependent, boolean horizontal) {
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
		cell.setFormulaValue(exprAfter.getFormulaString());
		
		System.out.println(">>>>"+expr+" extend to "+exprAfter.getFormulaString());
	}	

	public void shrink(Set<Ref> dependents, boolean horizontal) {
		for (Ref dependent : dependents) {
			System.out.println(">>>Shrink Sheet Formula: "+dependent);
			if (dependent.getType() == RefType.CELL) {
				shrinkCellRef(dependent,horizontal);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					shrinkChartRef((ObjectRef)dependent,horizontal);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					shrinkDataValidationRef((ObjectRef)dependent,horizontal);
				}
			} else {// TODO another

			}
		}
	}
	private void shrinkChartRef(ObjectRef dependent, boolean horizontal) {
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
	private void shrinkDataValidationRef(ObjectRef dependent, boolean horizontal) {
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

	private void shrinkCellRef(Ref dependent, boolean horizontal) {
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
		cell.setFormulaValue(exprAfter.getFormulaString());
		
		System.out.println(">>>>"+expr+" shrink to "+exprAfter.getFormulaString());
	}	
}
