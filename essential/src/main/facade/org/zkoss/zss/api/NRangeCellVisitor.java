package org.zkoss.zss.api;


public abstract class NRangeCellVisitor {

	public boolean createIfNotExist(int row, int column) {
		return true;
	}

	abstract public void visit(NRange cellRange);
}
