/* PostImport.java

	Purpose:
		
	Description:
		
	History:
		Nov 17, 2016 2:36:19 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.api;

import org.zkoss.zss.api.model.Book;

/**
 * Used with {@link Importer} for post-import-processing life cycle.
 * @author Henri
 * @since 3.9.1
 */
public interface PostImport {
	/**
	 * Used with {@link Importer} for post-import-processing life cycle.
	 * 
	 * After Importer.imports the book, the Importer will call back to this
	 * method with the loaded {@link Book} and you can use whatever Range
	 * methods to post-process the loaded Book.
	 * 
	 * Note that in this life cycle, all Range methods called will ignore
	 * autoRefresh(i.e. autoRefresh is always false) and skip all cell cache
	 * clearing; thus speed up the post processing.
	 * 
	 * @param book
	 */
	void process(Book book);
}
