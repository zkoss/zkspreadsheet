package org.zkoss.zss.ui;

import java.util.Map;

import org.zkoss.zss.api.model.Sheet;

/**
 * Action Handler for user's {@code Action}
 * 
 * @author dennis
 * 
 */
public interface UserActionHandler {

	void handleAction(Spreadsheet spreadsheet, Sheet targetSheet,String action,
			Rect selection, Map<String, Object> extraData);
}
