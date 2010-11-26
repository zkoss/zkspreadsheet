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

/**
 * @author Sam
 *
 */
public interface WorkbenchCtrl {

	public void openInsertFormulaDialog();
	
	public void openExportPdfDialog();
	
	public void openPasteSpecialDialog();
	
	public void openModifyRowHeightDialog();
	
	public void openCustomSortDialog();
	
	public void openHyperlinkDialog();
	
	public void toggleFormulaBar();
	
	public void openComposeFormulaDialog(FormulaMetaInfo metainfo);
	
	public void updateGridlinesCheckbox();

}
