package org.zkoss.zss.ngmodel;
/**
 * Indicate the object might has formula content
 * @author dennis
 *
 */
public interface FormulaContent {

	/**
	 * Clear the formula result cache if there is evaluation result
	 */
	public void clearFormulaResultCache();
	
	/**
	 * @return true if has parsing error, false if no error or not a formula content
	 */
	public boolean isFormulaParsingError();
	
}