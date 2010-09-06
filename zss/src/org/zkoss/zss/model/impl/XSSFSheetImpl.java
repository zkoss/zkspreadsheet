/* XSSFSheetImpl.java

	Purpose:
		
	Description:
		
	History:
		Sep 3, 2010 11:07:42 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.zss.model.Book;

/**
 * Implementation of {@link Sheet} based on XSSFSheet.
 * @author henrichen
 *
 */
public class XSSFSheetImpl extends XSSFSheet {
	//--XSSFSheet--//
    public XSSFSheetImpl() {
        super();
    }

    /**
     * Creates an XSSFSheet representing the given package part and relationship.
     * Should only be called by XSSFWorkbook when reading in an exisiting file.
     *
     * @param part - The package part that holds xml data represenring this sheet.
     * @param rel - the relationship of the given package part in the underlying OPC package
     */
    public XSSFSheetImpl(PackagePart part, PackageRelationship rel) {
        super(part, rel);
    }

	public Book getBook() {
		return (Book) getWorkbook();
	}
}
