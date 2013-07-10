package org.zkoss.zss.app.ui.dlg;

import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Map;

import org.zkoss.lang.Strings;
import org.zkoss.util.logging.Log;
import org.zkoss.util.media.Media;
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
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;


public class OpenBookCtrl extends DlgCtrlBase{
	private static final long serialVersionUID = 1L;
	private static final Log log = Log.lookup(OpenBookCtrl.class); 
	@Wire
	Listbox bookList;
	@Wire
	Button open;
	@Wire
	Button delete;
	
	ListModelList<Map<String,Object>> bookListModel = new ListModelList<Map<String,Object>>();
	
	
	@Override
	public void doAfterCompose(Window comp) throws Exception {
		super.doAfterCompose(comp);
		reloadBookModel();
	}
	
	private void reloadBookModel(){
		bookListModel = new ListModelList<Map<String,Object>>();
		BookRepository rep = getRepository();
		DateFormat dateFormat = DateFormat.getDateInstance(DateFormat.LONG);
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
		postCallback(DlgEvts.ON_OPEN, newMap(newEntry("bookinfo", selection.get("bookinfo"))));
		detach();
	}
	
	@Listen("onSelect=#bookList")
	public void onSelect(){
		updateButtonState();
	}
	
	private void updateButtonState(){
		boolean selected = bookListModel.getSelection().size()>0; 
		open.setDisabled(!selected);
		delete.setDisabled(!selected);
	}
	
	@Listen("onClick=#delete")
	public void onDelete(){
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
			e.printStackTrace();
			UiUtil.showInfoMessage("Can't delete book "+bookinfo.getName());
		}
	}
	
	@Listen("onClick=#cancel;onCancel=#openBookDlg")
	public void onCancel(){
		detach();
	}
	
	@Listen("onClick=#upload")
	public void onUpload(){
		Fileupload.get(5,new EventListener<UploadEvent>() {
			public void onEvent(UploadEvent event) throws Exception {
				BookInfo bookInfo = null;
				Book book = null;
				int count = 0;
				Importer importer = Importers.getImporter();
				BookRepository rep = getRepository();
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
							log.debug("exception when handling user upload file");
						}finally{
							if(is!=null){
								is.close();
							}
						}
					}
				}
				if(count==1){//open book directly if only one book
					postCallback(DlgEvts.ON_OPEN, newMap(newEntry("bookinfo", bookInfo),newEntry("book", book)));
					detach();
				}else if(count>0){
					reloadBookModel();
				}else{
					UiUtil.showInfoMessage("Doesn't get any supported files");
				}
			}
		});
	}
}
