package org.zkoss.zss.api;


public interface CellVisitor {

	public boolean createIfNotExist(int row, int column);

	public void visit(Range cellRange);
}
