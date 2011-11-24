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

import org.zkoss.poi.ss.usermodel.charts.ChartType;
import org.zkoss.zss.app.formula.FormulaMetaInfo;

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
	
	public void openCustomSortDialog();
	
	public void openHyperlinkDialog();
	
	public boolean toggleFormulaBar();
	
	public void openComposeFormulaDialog(FormulaMetaInfo metainfo);
	
	public void openFormatNumberDialog();
	
	public void openSaveFileDialog();
		
	public void updateGridlinesCheckbox();

	/**
	 * 
	 * @param headerType header type {@link WorkbookCtrl}
	 */
	public void openModifyHeaderSizeDialog(int headerType);
	
	/**
	 * Open rename sheet dialog
	 */
	public void openRenameSheetDialog(String originalSheetName);
	
	public void openAutoFilterDialog(Object data);
}
