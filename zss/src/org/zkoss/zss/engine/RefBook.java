/* RefBook.java

	Purpose:
		
	Description:
		
	History:
		Mar 7, 2010 5:35:25 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine;

import java.util.Set;

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;

/**
 * Internal Use Only. Book that handle the {@link RefSheet}s.
 * @author henrichen
 *
 */
public interface RefBook {
	/**
	 * Returns the book name of this reference book. 
	 * @return the book name of this reference book.
	 */
	public String getBookName();
	
	/**
	 * Returns (or create then return if not exist) the {@link RefSheet}  
	 * at specified index.
	 * @param sheetname the name of the sheet to be gotten/created
	 * @return the {@link RefSheet} for the data sheet at specified index.
	 */
	public RefSheet getOrCreateRefSheet(String sheetname);
	
	/**
	 * Returns the {@link RefSheet} of the specified sheet name; return null if not exists.
	 * @param sheetname the name of the sheet to be gotten
	 * @return the {@link RefSheet} of the specified sheet name; return null if not exists.
	 */
	public RefSheet getRefSheet(String sheetname);
	
	/**
	 * Remove and return the {@link RefSheet} for the data sheet of the  
	 * specified sheet name; return null if not exists.
	 * @param sheetname the name of the sheet to be removed
	 * @return the removed RefSheet; return null if not exists.
	 */
	public RefSheet removeRefSheet(String sheetname);
	
	/**
	 * Update the sheet name of the specified sheet to new sheet name.
	 * @param oldsheetname old sheet name used to located the sheet
	 * @param newsheetname new sheet name to be set into the located sheet
	 */
	public void setSheetName(String oldsheetname, String newsheetname);
	
	/**
	 * Returns the maximum row index of this reference book 
	 * (e.g. 64K-1 is the maximum row index of Excel 97, 1M-1 is of Excel 2007). 
	 * @return The maximum row index of this reference book 
	 */
	public int getMaxrow();
	
	/**
	 * Returns the maximum column index of this refernece book 
	 * (e.g. 255 is the maximum column index of Excel 97, 16K-1 is of Excel 2007). 
	 * @return the maximum column index of this refernece book
	 */
	public int getMaxcol();

	/**
	 * Subscribe a event listener to this reference book.
	 * @param listener the event listener that will handle event fired by this reference book.
	 */
	public void subscribe(EventListener listener);

	/**
	 * Un-subscribe the event listener from this reference book.
	 * @param listener the event listener that will handle event fired by this reference book.
	 */
	public void unsubscribe(EventListener listener);
	
	/**
	 * Publish an event to this reference book.
	 * @param event the event to be fired by this reference book.
	 */
	public void publish(Event event);
	
	/**
	 * Return or create if not exist the specified variable reference.
	 * @param name the name of the specified variable reference
	 * @param dummy dummy sheet, used to locate associated RefBook 
	 * @return the variable reference.
	 */
	public Ref getOrCreateVariableRef(String name, RefSheet dummy);
	
	/**
	 * Remove the specified variable reference. 
	 * @param name the name of the specified variable reference
	 * @return the removed variable reference
	 */
	public Ref removeVariableRef(String name);
	
	/**
	 * Returns both all and last dependent cell references that are affected by the specified variable.
	 * @param name the variable name
	 * @return reference set: [0] the "last" dependent cell references to be re-evaluated the associated cell value; 
	 * 	[1] the "all" dependent cell references to be reloaded the associated cell value.
	 */
	public Set<Ref>[] getBothDependents(String name);

	/**
	 * Sets share scope of this reference book.
	 * @param scope share scope of thie reference book.
	 */
	public void setShareScope(String scope);
	
	/**
	 * Returns share scope of this reference book.
	 * @return share scope of this reference book.
	 */
	public String getShareScope();
}