/* Book.java

	Purpose:
		
	Description:
		
	History:
		Mar 22, 2010 7:11:11 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.sys;

import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.FormulaEvaluator;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.ss.usermodel.PictureData;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.EventQueues;

/**
 * ZK Spreadsheet book.
 * @author henrichen
 *
 */
public interface XBook extends Workbook {
	/**
	 * Returns the spreadsheet version of this book (EXCEL97 or EXCEL2007).
	 * @return the spreadsheet version of this book.
	 */
	public SpreadsheetVersion getSpreadsheetVersion();
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
	public Font getDefaultFont(); 	//http://support.microsoft.com/kb/214123

	/**
	 * Sets the default font used in this book.
	 * @param font the font to be used as the default font in this book.
	 */
	public void setDefaultFont(Font font);
	
	/**
	 * Returns default character width in pixel.
	 * @return default character width in pixel.
	 */
	public int getDefaultCharWidth();
	
	/**
	 * Notify this book the value of the specified variable has changed. 
	 * @param variables the variable names
	 */
	public void notifyChange(String[] variables);
	
	/**
	 * Returns the repeat rows and columns in a CellRangeAddress per the specified sheet; -1 mean no repeat rows or columns
	 * @param sheetIndex the sheet index
	 * @return the repeat rows and columns in a CellRangeAddress; 
	 */
	public CellRangeAddress getRepeatingRowsAndColumns(int sheetIndex);
	
    /**
     * Finds a font that matches the one with the supplied attributes
     *
     * @return the font with the matched attributes or <code>null</code>
     */
	public Font findFont(short boldWeight, Color color, short fontHeight, String name, boolean italic, boolean strikeout, short typeOffset, byte underline);
	
	/**
	 * Sets share scope of this book; default scope is {@link EventQueues#DESKTOP}.
	 * <p>Note: this feature requires ZK Spreadsheet EE.</p>
	 * @param scope share scope of this book: can be {@link EventQueues#DESKTOP},{@link EventQueues#GROUP},{@link EventQueues#SESSION},{@link EventQueues#APPLICATION}. 
	 */
	public void setShareScope(String scope);
	
	/**
	 * Returns share scope of this book.
	 * <p>Note: this feature requires ZK Spreadsheet EE.</p>
	 * @see #setShareScope(String)
	 */
	public String getShareScope();
	
	/**
	 * Returns ZK Spreadsheet {@link XSheet} by name.
	 */
	public XSheet getWorksheet(String name);
	
	/**
	 * Returns ZK Spreadsheet {@link XSheet} by index(0-based). 
	 */
	public XSheet getWorksheetAt(int index);

    /**
     * Delete the PictureData.
     * @param pictureData
     */
    void deletePictureData(PictureData pictureData);
    
    /**
     * Gets a boolean value that indicates whether the date systems used in the workbook starts in 1904.
     * <p>
     * The default value is false, meaning that the workbook uses the 1900 date system,
     * where 1/1/1900 is the first day in the system..
     * </p>
     * @return true if the date systems used in the workbook starts in 1904
     */
    boolean isDate1904();
}
