/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ui.dlg;

import java.io.IOException;
import java.io.InputStream;
import java.io.Serializable;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;

import org.zkoss.util.logging.Log;
import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.app.repository.BookInfo;
import org.zkoss.zss.app.repository.BookRepository;
import org.zkoss.zss.app.repository.BookRepositoryFactory;
import org.zkoss.zss.app.ui.UiUtil;
import org.zkoss.zul.Button;
import org.zkoss.zul.Fileupload;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Window;
import org.zkoss.zul.event.ListDataEvent;

/**
 * 
 * @author dennis
 *
 */
public class OpenManageBookCtrl extends DlgCtrlBase{
	private static final long serialVersionUID = 1L;
	
	public final static String ARG_BOOKINFO = "bookinfo";
	public final static String ARG_BOOK = "book";
	
	private final static String URI = "~./zssapp/dlg/openManageBook.zul";
	
	public static final String ON_OPEN = "onOpen";
	
	private static final Log log = Log.lookup(OpenManageBookCtrl.class); 
	@Wire
	Listbox bookList;
	@Wire
	Button open;
	@Wire
	Button delete;
	@Wire
	Button upload;
	
	ListModelList<Map<String,Object>> bookListModel = new ListModelList<Map<String,Object>>();
	
	public static void show(EventListener<DlgCallbackEvent> callback) {
		Map arg = newArg(callback);
		
		Window comp = (Window)Executions.createComponents(URI, null, arg);
		comp.doModal();
		return;
	}
	
	@Override
	public void doAfterCompose(Window comp) throws Exception {
		super.doAfterCompose(comp);

		reloadBookModel();
	}
	
	private void reloadBookModel(){
		bookListModel = new ListModelList<Map<String,Object>>();
		BookRepository rep = getRepository();
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
		for(BookInfo info : rep.list()){
			Map<String,Object> data = new HashMap<String,Object>();
			data.put("name", info.getName());
			data.put("lastmodify", dateFormat.format(info.getLastModify()));
			data.put("bookinfo", info);
			bookListModel.add(data);
		}
		bookList.setModel(bookListModel);
		updateButtonState();
	}
	
	private BookRepository getRepository(){
		return BookRepositoryFactory.getInstance().getRepository();
	}
	
	@Listen("onClick=#open; onDoubleClick=#bookList")
	public void onOpen(){
		Map<String,Object> selection = (Map<String,Object>)UiUtil.getSingleSelection(bookListModel);
		if(selection==null){
			UiUtil.showInfoMessage("Please select a book first");
			return;
		}
		postCallback(ON_OPEN, newMap(newEntry(ARG_BOOKINFO, selection.get("bookinfo"))));
		detach();
	}
	
	@Listen("onSelect=#bookList")
	public void onSelect(){
		updateButtonState();
	}
	
	private void updateButtonState(){
		boolean selected = bookListModel.getSelection().size()>0;
		
		boolean readonly = UiUtil.isRepositoryReadonly();

		open.setDisabled(!selected);
		delete.setDisabled(!selected || readonly);
		upload.setDisabled(readonly);
	}
	
	@Listen("onClick=#delete")
	public void onDelete(){
		if(UiUtil.isRepositoryReadonly()){
			return;
		}
		Map<String,Object> selection = (Map<String,Object>)UiUtil.getSingleSelection(bookListModel);
		if(selection==null){
			UiUtil.showInfoMessage("Please select a book first");
			return;
		}
		
		BookInfo bookinfo = (BookInfo)selection.get("bookinfo");
		BookRepository rep = getRepository();
		try {
			rep.delete(bookinfo);
			reloadBookModel();
		} catch (IOException e) {
			log.error(e.getMessage(),e);
			UiUtil.showInfoMessage("Can't delete book "+bookinfo.getName());
		}
	}
	
	@Listen("onClick=#cancel;onCancel=#openBookDlg")
	public void onCancel(){
		detach();
	}
	
	@Listen("onClick=#upload")
	public void onUpload(){
		if(UiUtil.isRepositoryReadonly()){
			return;
		}
		Fileupload.get(5,new EventListener<UploadEvent>() {
			public void onEvent(UploadEvent event) throws Exception {
				BookInfo bookInfo = null;
				Book book = null;
				int count = 0;
				Importer importer = Importers.getImporter();
				BookRepository rep = getRepository();
				Media[] medias = event.getMedias();
				if(medias==null)
					return;
				for(Media m:event.getMedias()){
					if(m.isBinary()){
						InputStream is = null;
						try{
							is = m.getStreamData();
							String name = m.getName();						
							book = importer.imports(is, name);
							bookInfo = rep.saveAs(name, book);
							count++;
						}catch(Exception x){
							log.debug(x);
							log.warning("exception when handling user upload file", x);
						}finally{
							if(is!=null){
								is.close();
							}
						}
					}
				}
				if(count==1){//open book directly if only one book
					postCallback(ON_OPEN, newMap(newEntry(ARG_BOOKINFO, bookInfo),newEntry(ARG_BOOK, book)));
					detach();
				}else if(count>0){
					reloadBookModel();
				}else{
					UiUtil.showInfoMessage("Can't get any supported files");
				}
			}
		});
	}
	
	static private class MapAttrComparator implements Comparator<Map<String, Object>>, Serializable {
		private static final long serialVersionUID = -2889285276678010655L;
		private final boolean _asc;
		private final String _attr;
		
		public MapAttrComparator(boolean asc,String attr) {
			_asc = asc;
			_attr = attr;
		}

		@Override
		public int compare(Map<String, Object> o1, Map<String, Object> o2) {
			if(_asc){
				return o1.get(_attr).toString().compareTo(o2.get(_attr).toString());
			}else{
				return o2.get(_attr).toString().compareTo(o1.get(_attr).toString());
			}
		}
	}
	
	public Comparator<Map<String, Object>> getBookNameDescComparator() {
		return new MapAttrComparator(false, "name");
	}

	public Comparator<Map<String, Object>> getBookNameAscComparator() {
		return new MapAttrComparator(true, "name");
	}

	public Comparator<Map<String, Object>> getBookDateDescComparator() {
		return new MapAttrComparator(false, "lastmodify");
	}

	public Comparator<Map<String, Object>> getBookDateAscComparator() {
		return new MapAttrComparator(true, "lastmodify");
	}
	
}
