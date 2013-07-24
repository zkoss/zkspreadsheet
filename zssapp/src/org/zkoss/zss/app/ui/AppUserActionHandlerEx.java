/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ui;

import java.util.Set;

import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.DefaultUserAction;
import org.zkoss.zssex.ui.DefaultUserActionHandlerEx;
/**
 * Application handler for EE
 * @author dennis
 *
 */
public class AppUserActionHandlerEx extends DefaultUserActionHandlerEx  {
		private static final long serialVersionUID = 1L;
		AppCtrl appCtrl;
		
		public AppUserActionHandlerEx(AppCtrl appCtrl){
			this.appCtrl = appCtrl;
		}
		
		
		@Override
		public Set<String> getSupportedUserAction(Sheet sheet) {
			Set<String> actions = super.getSupportedUserAction(sheet);
			actions.add(DefaultUserAction.NEW_BOOK.getAction());
			if(sheet!=null){
				boolean readonly = UiUtil.isRepositoryReadonly();
				if(!readonly){
					actions.add(DefaultUserAction.SAVE_BOOK.getAction());
				}
				actions.add(DefaultUserAction.EXPORT_PDF.getAction());
			}
			return actions;
		}
		
		@Override
		protected boolean dispatchAction(String action) {
			DefaultUserAction dua = DefaultUserAction.getBy(action);

			if (DefaultUserAction.NEW_BOOK.equals(dua)) {
				appCtrl.doOpenNewBook();
				return true;
			} else if (DefaultUserAction.SAVE_BOOK.equals(dua)) {
				appCtrl.doSaveBook(false);
				return true;
			} else if (DefaultUserAction.EXPORT_PDF.equals(dua)) {
				appCtrl.doExportPdf();
				return true;
			} 

			return super.dispatchAction(action);
		}
		
		@Override
		protected boolean doCloseBook(){
			super.doCloseBook();
			appCtrl.doCloseBook();
			return true;
		}
		
		protected boolean doSheetSelect() {
			super.doSheetSelect();
			appCtrl.doSheetSelect();
			return true;
		}
	}