/* WorkspaceContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 23, 2010 5:43:38 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zss.app.file.FileHelper;
import org.zkoss.zss.app.file.SpreadSheetMetaInfo;
import org.zkoss.zss.app.file.UnsupportedSpreadSheetFileException;

/**
 * @author Sam
 *
 */
public class WorkspaceContext extends AbstractBaseContext {
	
	SpreadSheetMetaInfo info;
	
	public static WorkspaceContext getInstance(Desktop desktop) {
		WorkspaceContext context = (WorkspaceContext)desktop.getAttribute("WorkspaceContext");
		if (context == null) {
			desktop.setAttribute("WorkspaceContext", context = new WorkspaceContext());
		}
		return context;
	}
	
	public List<SpreadSheetMetaInfo> getMetainfos() {
		Map<String, SpreadSheetMetaInfo> metaInfos = SpreadSheetMetaInfo.getMetaInfos();
		return new ArrayList<SpreadSheetMetaInfo>(metaInfos.values());
	}
	
	public SpreadSheetMetaInfo getCurrent() {
		return info;
	}
	
	public SpreadSheetMetaInfo store(Media media) throws UnsupportedSpreadSheetFileException {
		return FileHelper.store(media);
	}
	
	public void delete(String src) {
		FileHelper.deleteSpreadsheet(src);
	}
}