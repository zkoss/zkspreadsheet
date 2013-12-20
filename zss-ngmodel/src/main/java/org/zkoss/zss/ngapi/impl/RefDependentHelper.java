package org.zkoss.zss.ngapi.impl;

import java.util.HashSet;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NChart;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef.ObjectType;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;

/*package*/ class RefDependentHelper {
	final NBookSeries bookSeries;
	public RefDependentHelper(NBookSeries bookSeries) {
		this.bookSeries = bookSeries;
	}

	public void handle(HashSet<Ref> dependentSet) {
		// clear formula cache
		for (Ref dependent : dependentSet) {
			System.out.println(">>> Dependent "+dependent);
			//clear the dependent's formula cache since the precedent is changed.
			if (dependent.getType() == RefType.CELL) {
				handleCellRef(dependent);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					handleChartRef((ObjectRef)dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.VALIDATION){
					
				}
			} else {// TODO another

			}
		}
	}
	private void handleChartRef(ObjectRef dependent) {
		NBook book = bookSeries.getBook(dependent
				.getBookName());
		NSheet sheet = book.getSheetByName(dependent
				.getSheetName());
		String[] ids = dependent.getObjectIdPath();
		NChart chart = sheet.getChart(ids[0]);
		chart.getData().clearFormulaResultCache();
	}

	private void handleCellRef(Ref dependent) {
		NBook book = bookSeries.getBook(dependent
				.getBookName());
		NSheet sheet = book.getSheetByName(dependent
				.getSheetName());
		NCell cell = sheet.getCell(dependent.getRow(),
				dependent.getColumn());
		cell.clearFormulaResultCache();
	}
}
