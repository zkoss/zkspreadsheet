/* WorkbenchCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 25, 2010 2:10:26 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import org.zkoss.zss.app.formula.FormulaMetaInfo;
import org.zkoss.zss.ui.Rect;

/**
 * @author Sam
 *
 */
public interface WorkbenchCtrl {

	public void openInsertFormulaDialog();
	
	/**
	 * Open export pdf dialog
	 */
	public void openExportPdfDialog();
	
	public void openExportHtmlDialog();
	
	public void openOpenFileDialog();
	
	public void openImportFileDialog();
	
	public void openPasteSpecialDialog();
	
	public void openCustomSortDialog(Rect selection);
	
	public void openHyperlinkDialog(Rect selection);
	
	public boolean toggleFormulaBar();
	
	public void openComposeFormulaDialog(FormulaMetaInfo metainfo);
	
	public void openFormatNumberDialog(Rect selection);
	
	public void openSaveFileDialog();

	/**
	 * 
	 * @param headerType header type {@link WorkbookCtrl}
	 */
	public void openModifyHeaderSizeDialog(int headerType, Rect selection);
	
	/**
	 * Open rename sheet dialog
	 */
	public void openRenameSheetDialog(String originalSheetName);
}
