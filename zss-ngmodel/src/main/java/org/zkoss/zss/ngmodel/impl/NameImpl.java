package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

public class NameImpl extends NameAdv {

	private final String id;
	private BookAdv book;
	private String name;
	
	private String refersToExpr;
	
	private CellRegion refersTo;
	private String sheetName;
	
	public NameImpl(BookAdv book, String id) {
		this.book = book;
		this.id = id;
	}

	@Override
	public String getName() {
		return name;
	}

	@Override
	public String getSheetName() {
		return sheetName;
	}

	@Override
	public CellRegion getRefersTo() {
		return refersTo;
	}

	@Override
	public String getRefersToFormula() {
		return refersToExpr;
	}

	@Override
	public void release() {
		checkOrphan();
		clearFormulaDependency();
		book = null;
	}

	@Override
	public void checkOrphan() {
		if(book==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	void setName(String newname) {
		name = newname;
	}

	@Override
	public String getId() {
		return id;
	}

	@Override
	public void setRefersToFormula(String refersToExpr) {
		checkOrphan();
		this.refersToExpr = refersToExpr;
		//TODO support function as Excel (POI)
		
		
		
		
		//TODO use formula engine to keep dependency info
	}

	@Override
	public void clearFormulaResultCache() {
		// TODO Auto-generated method stub
	}

	private void clearFormulaDependency() {
		Ref ref = new RefImpl(this);
		((BookSeriesAdv)book.getBookSeries()).getDependencyTable().clearDependents(ref);
	}

	@Override
	BookAdv getBook() {
		checkOrphan();
		return book;
	}
}
