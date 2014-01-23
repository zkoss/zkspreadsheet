package org.zkoss.zss.ngmodel.impl;

import java.util.Set;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NChart;
import org.zkoss.zss.ngmodel.NDataValidation;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.SheetRegion;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef.ObjectType;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;

/*package*/ class FormulaShiftHelper {
	final NBookSeries bookSeries;
	final SheetRegion sheetRegion;

	public FormulaShiftHelper(AbstractBookSeriesAdv bookSeries,
			SheetRegion sheetRegion) {
		this.bookSeries = bookSeries;
		this.sheetRegion = sheetRegion;
	}

	public void shift(Set<Ref> dependents,int rowOffset, int columnOffset) {
		for (Ref dependent : dependents) {
			System.out.println(">>>Shift Sheet Formula: "+dependent);
			if (dependent.getType() == RefType.CELL) {
				shiftCellRef(dependent,rowOffset,columnOffset);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					shiftChartRef((ObjectRef)dependent,rowOffset,columnOffset);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					shiftDataValidationRef((ObjectRef)dependent,rowOffset,columnOffset);
				}
			} else {// TODO another

			}
		}
	}
	private void shiftChartRef(ObjectRef dependent,int rowOffset, int columnOffset) {
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
	private void shiftDataValidationRef(ObjectRef dependent,int rowOffset, int columnOffset) {
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

	private void shiftCellRef(Ref dependent,int rowOffset, int columnOffset) {
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
		FormulaExpression exprAfter = engine.shift(expr, sheetRegion, rowOffset, columnOffset, new FormulaParseContext(sheet, null));//null ref, no trace dependence here
		cell.setFormulaValue(exprAfter.getFormulaString());
		
		System.out.println(">>>>"+expr+" shift to "+exprAfter.getFormulaString());
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
