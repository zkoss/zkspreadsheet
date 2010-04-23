/* Book.java

	Purpose:
		
	Description:
		
	History:
		Mar 22, 2010 7:11:11 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.engine.EventDispatcher;

/**
 * Spreadsheet book.
 * @author henrichen
 *
 */
public interface Book extends Workbook {
	/**
	 * Returns the associated book name of this spreadsheet book (important when used with @{link Books}).
	 * @return the associated book name of this spreadsheet book.
	 */
	public String getBookName();
	
	/**
	 * Returns the associated formula evaluator for this spreadsheet book.
	 * @return the associated formula evaluator for this spreadsheet book.
	 */
	public FormulaEvaluator getFormulaEvaluator();
	
	/** Adds a variable resolver to this model.
	 * The first added variable resolver has the highest priority.
	 * @param resolver the variable resolver to be added.
	 */
	public void addVariableResolver(VariableResolver resolver);
	/** Removes a variable resolver to this model.
	 * @param resolver the variable resolver to be removed.
	 */
	public void removeVariableResolver(VariableResolver resolver);

	/** Adds a function mapper to this model for the specified prefix.
	 * The first added function mapper has the highest priority.
	 *
	 * @param mapper the function mapper
	 */
	public void addFunctionMapper(FunctionMapper mapper);
	/** Removes a function mapper to this model for the specified prefix.
	 *
	 * @param mapper the function mapper.prefix the prefix of the specified mapper.
	 */
	public void removeFunctionMapper(FunctionMapper mapper);
	
	/**
	 * Subscribe a listener to listener to this book
	 * @param listener 
	 */
	public void subscribe(EventListener listener);
	
	/**
	 * Unsubscribe the specified listener from listening to this book
	 * @param listener
	 */
	public void unsubscribe(EventListener listener);
	
	/**
	 * Returns the default font used in this book.
	 * @return the default font used in this book.
	 */
	public Font getDefaultFont();
	
	/**
	 * Notify this book the value of the specified variable has changed. 
	 * @param variables the variable names
	 */
	public void notifyChange(String[] variables);
}
