package org.zkoss.zss.model.impl;

import java.util.Set;

import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeries;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SChart;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.sys.dependency.ObjectRef;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.dependency.ObjectRef.ObjectType;
import org.zkoss.zss.model.sys.dependency.Ref.RefType;

/*package*/ class FormulaCacheClearHelper {
	final SBookSeries bookSeries;
	public FormulaCacheClearHelper(SBookSeries bookSeries) {
		this.bookSeries = bookSeries;
	}

	public void clear(Set<Ref> dependents) {
		// clear formula cache
		for (Ref dependent : dependents) {
			System.out.println(">>>Clear Formula Cache: "+dependent);
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
		SBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		String[] ids = dependent.getObjectIdPath();
		SChart chart = sheet.getChart(ids[0]);
		if(chart!=null){
			chart.getData().clearFormulaResultCache();
		}
	}
	private void handleDataValidationRef(ObjectRef dependent) {
		SBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		String[] ids = dependent.getObjectIdPath();
		SDataValidation validation = sheet.getDataValidation(ids[0]);
		if(validation!=null){
			validation.clearFormulaResultCache();
		}
	}

	private void handleCellRef(Ref dependent) {
		SBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		SCell cell = sheet.getCell(dependent.getRow(),
				dependent.getColumn());
		cell.clearFormulaResultCache();
	}
}
