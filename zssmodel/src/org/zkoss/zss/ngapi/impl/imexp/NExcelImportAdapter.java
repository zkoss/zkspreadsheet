/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by Hawk
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl.imexp;

import java.io.*;

import org.zkoss.poi.POIXMLDocument;
import org.zkoss.poi.poifs.filesystem.POIFSFileSystem;
import org.zkoss.zss.model.SBook;
/**
 * 
 * @author dennis
 * @author Hawk
 * @since 3.5.0
 */
public class NExcelImportAdapter extends AbstractImporter{

	@Override
	public SBook imports(InputStream is, String bookName) throws IOException {
		if(!is.markSupported()) {
			is = new PushbackInputStream(is, 8);
		}
		if (POIFSFileSystem.hasPOIFSHeader(is)) {
			return new NExcelXlsImporter().imports(is, bookName);
		}else if (POIXMLDocument.hasOOXMLHeader(is)) {
			return new NExcelXlsxImporter().imports(is, bookName);
		}
		throw new IllegalArgumentException("The input stream to be imported is neither an OLE2 stream, nor an OOXML stream");
	}
}
