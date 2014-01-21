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
	final int rowOffset;
	final int columnOffset;

	public FormulaShiftHelper(AbstractBookSeriesAdv bookSeries,
			SheetRegion sheetRegion, int rowOffset, int columnOffset) {
		this.bookSeries = bookSeries;
		this.sheetRegion = sheetRegion;
		this.rowOffset = rowOffset;
		this.columnOffset = columnOffset;
	}

	public void shift(Set<Ref> dependents) {
		// clear formula cache
		for (Ref dependent : dependents) {
			System.out.println(">>>Sheet Formula: "+dependent);
			//clear the dependent's formula cache since the precedent is changed.
			if (dependent.getType() == RefType.CELL) {
				handleCellRef(dependent);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					handleChartRef((ObjectRef)dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.DATA_VALIDATION){
					handleDataValidationRef((ObjectRef)dependent);
				}
			} else {// TODO another

			}
		}
	}
	private void handleChartRef(ObjectRef dependent) {
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
	private void handleDataValidationRef(ObjectRef dependent) {
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

	private void handleCellRef(Ref dependent) {
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
}
