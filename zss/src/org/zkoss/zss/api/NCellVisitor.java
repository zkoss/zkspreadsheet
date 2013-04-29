package org.zkoss.zss.api;


public interface NCellVisitor {

	public boolean createIfNotExist(int row, int column);

	public void visit(NRange cellRange);
}
