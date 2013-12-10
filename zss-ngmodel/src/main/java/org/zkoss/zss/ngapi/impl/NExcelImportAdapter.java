/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl;

import java.io.IOException;
import java.io.InputStream;

import org.zkoss.poi.POIXMLDocument;
import org.zkoss.poi.poifs.filesystem.POIFSFileSystem;
import org.zkoss.zss.ngmodel.NBook;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NExcelImportAdapter extends AbstractImporter{

	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {
		if (POIFSFileSystem.hasPOIFSHeader(is)) {
			return new NExcelXlsImporter().imports(is, bookName);
		}else if (POIXMLDocument.hasOOXMLHeader(is)) {
			return new NExcelXlsxImporter().imports(is, bookName);
		}
		throw new IllegalArgumentException("The input stream to be imported is neither an OLE2 stream, nor an OOXML stream");
	}
}
