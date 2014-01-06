package org.zkoss.zss.ngmodel.impl;

import java.util.Set;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NChart;
import org.zkoss.zss.ngmodel.NDataValidation;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef.ObjectType;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;

/*package*/ class FormulaCacheClearHelper {
	final NBookSeries bookSeries;
	public FormulaCacheClearHelper(NBookSeries bookSeries) {
		this.bookSeries = bookSeries;
	}

	public void clear(Set<Ref> dependents) {
		// clear formula cache
		for (Ref dependent : dependents) {
			System.out.println(">>> Clear "+dependent);
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
		NBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		String[] ids = dependent.getObjectIdPath();
		NChart chart = sheet.getChart(ids[0]);
		if(chart!=null){
			chart.getData().clearFormulaResultCache();
		}
	}
	private void handleDataValidationRef(ObjectRef dependent) {
		NBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		String[] ids = dependent.getObjectIdPath();
		NDataValidation validation = sheet.getDataValidation(ids[0]);
		if(validation!=null){
			validation.clearFormulaResultCache();
		}
	}

	private void handleCellRef(Ref dependent) {
		NBook book = bookSeries.getBook(dependent.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(dependent.getSheetName());
		if(sheet==null) return;
		NCell cell = sheet.getCell(dependent.getRow(),
				dependent.getColumn());
		cell.clearFormulaResultCache();
	}
}
