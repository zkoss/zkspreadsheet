package org.zkoss.zss;

import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.poi.xssf.usermodel.XSSFComment;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;

public class POIUtil {

	/**
	 * @return the comment if there is a comment at given position; null, otherwise.
	 */
	public static XSSFComment getComment(XSSFSheet sheet, String ref) {
		CellReference cr = new CellReference(ref);
		return sheet.getCellComment(cr.getRow(), cr.getCol());
	}
}
