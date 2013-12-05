package org.zkoss.zss.ngmodel;

public interface NName extends FormulaContent{

	public String getId();
	public String getName();
	
	public String getSheetName();
	public CellRegion getRefersTo();
	
	public String getRefersToFormula();
	public void setRefersToFormula(String refersExpr);
}
