///* 
//	Purpose:
//		
//	Description:
//		
//	History:
//		2013/7/10, Created by dennis
//
//Copyright (C) 2013 Potix Corporation. All Rights Reserved.
//
//*/
//package org.zkoss.zss.app.ui;
//
//import java.util.Set;
//
//import org.zkoss.zss.api.model.Sheet;
//import org.zkoss.zss.ui.sys.DefaultAuxAction;
//import org.zkoss.zss.ui.sys.DefaultComponentActionManager;
//import org.zkoss.zss.ui.sys.DefaultComponentActionManagerX;
///**
// * Application handler for ose
// * @author dennis
// *
// */
//public class AppComponentActionManager extends DefaultComponentActionManagerX  {
//		private static final long serialVersionUID = 1L;
//		AppCtrl appCtrl;
//		
//		public AppComponentActionManager(AppCtrl appCtrl){
//			this.appCtrl = appCtrl;
//		}
//		
//		
//		@Override
//		public Set<String> getSupportedUserAction(Sheet sheet) {
//			Set<String> actions = super.getSupportedUserAction(sheet);
//			actions.add(DefaultAuxAction.NEW_BOOK.getAction());
//			if(sheet!=null){
//				boolean readonly = UiUtil.isRepositoryReadonly();
//				if(!readonly){
//					actions.add(DefaultAuxAction.SAVE_BOOK.getAction());
//				}
//				actions.add(DefaultAuxAction.EXPORT_PDF.getAction());
//			}
//			return actions;
//		}
//		
//		@Override
//		protected boolean dispatchAction(String action) {
//			DefaultAuxAction dua = DefaultAuxAction.getBy(action);
//
//			if (DefaultAuxAction.NEW_BOOK.equals(dua)) {
//				appCtrl.doOpenNewBook();
//				return true;
//			} else if (DefaultAuxAction.SAVE_BOOK.equals(dua)) {
//				appCtrl.doSaveBook(false);
//				return true;
//			} else if (DefaultAuxAction.EXPORT_PDF.equals(dua)) {
//				appCtrl.doExportPdf();
//				return true;
//			} 
//
//			return super.dispatchAction(action);
//		}
//		
//		@Override
//		protected boolean doCloseBook(){
//			super.doCloseBook();
//			appCtrl.doCloseBook();
//			return true;
//		}
//		
//		protected boolean doSheetSelect() {
//			super.doSheetSelect();
//			appCtrl.doSheetSelect();
//			return true;
//		}
//	}