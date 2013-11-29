package org.zkoss.zss.ngmodel;

public interface NDataValidationConstraint {

	public enum ConstraintType {
		ANY, INTEGER, DECIMAL, LIST, DATE, TIME, TEXT_LENGTH
		/*, FORMULA*/;

	}
	
	public ConstraintType getType();
}
