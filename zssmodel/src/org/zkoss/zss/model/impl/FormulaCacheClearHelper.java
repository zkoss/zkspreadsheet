package org.zkoss.zss.model.impl;

import java.util.Iterator;
import java.util.Set;

import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeries;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SChart;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SRow;
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

	public void clear(Set<Ref> refs) {
		// clear formula cache
		for (Ref ref : refs) {
			System.out.println(">>>Clear Formula Cache: "+ref);
			//clear the dependent's formula cache since the precedent is changed.
			if (ref.getType() == RefType.CELL || ref.getType() == RefType.AREA) {
				handleAreaRef(ref);
			} else if (ref.getType() == RefType.OBJECT) {
				if(((ObjectRef)ref).getObjectType()==ObjectType.CHART){
					handleChartRef((ObjectRef)ref);
				}else if(((ObjectRef)ref).getObjectType()==ObjectType.DATA_VALIDATION){
					handleDataValidationRef((ObjectRef)ref);
				}
			} else {// TODO another

			}
		}
	}
	private void handleChartRef(ObjectRef ref) {
		SBook book = bookSeries.getBook(ref.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(ref.getSheetName());
		if(sheet==null) return;
		String[] ids = ref.getObjectIdPath();
		SChart chart = sheet.getChart(ids[0]);
		if(chart!=null){
			chart.getData().clearFormulaResultCache();
		}
	}
	private void handleDataValidationRef(ObjectRef ref) {
		SBook book = bookSeries.getBook(ref.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(ref.getSheetName());
		if(sheet==null) return;
		String[] ids = ref.getObjectIdPath();
		SDataValidation validation = sheet.getDataValidation(ids[0]);
		if(validation!=null){
			validation.clearFormulaResultCache();
		}
	}

	private void handleAreaRef(Ref ref) {
		SBook book = bookSeries.getBook(ref.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(ref.getSheetName());
		if(sheet==null) return;
		
		for(int r = ref.getRow();r<=ref.getLastRow();r++){
			for(int c = ref.getColumn();c<=ref.getLastColumn();c++){
				SCell cell = ((AbstractSheetAdv)sheet).getCell(r,c,false);
				if(cell!=null){
					cell.clearFormulaResultCache();
				}
			}
		}
	}
}
