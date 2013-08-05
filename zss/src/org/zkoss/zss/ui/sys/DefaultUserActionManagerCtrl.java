/* DefaultComponentActionHandler.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/06/03 by Dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
 */
package org.zkoss.zss.ui.sys;

import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.zkoss.lang.Strings;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.CellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.ui.AuxAction;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.UserActionContext;
import org.zkoss.zss.ui.UserActionContext.Clipboard;
import org.zkoss.zss.ui.UserActionHandler;
import org.zkoss.zss.ui.UserActionManager;
import org.zkoss.zss.ui.event.AuxActionEvent;
import org.zkoss.zss.ui.event.CellSelectionAction;
import org.zkoss.zss.ui.event.CellSelectionUpdateEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.KeyEvent;
import org.zkoss.zss.ui.sys.ua.impl.AbstractBookHandler;
import org.zkoss.zss.ui.sys.ua.impl.AbstractProtectedHandler;
import org.zkoss.zss.ui.sys.ua.impl.AddSheetHandler;
import org.zkoss.zss.ui.sys.ua.impl.ApplyBorderHandler;
import org.zkoss.zss.ui.sys.ua.impl.ClearCellHandler;
import org.zkoss.zss.ui.sys.ua.impl.CloseBookHandler;
import org.zkoss.zss.ui.sys.ua.impl.CopyHandler;
import org.zkoss.zss.ui.sys.ua.impl.CutHandler;
import org.zkoss.zss.ui.sys.ua.impl.DeleteCellLeftHandler;
import org.zkoss.zss.ui.sys.ua.impl.DeleteCellUpHandler;
import org.zkoss.zss.ui.sys.ua.impl.DeleteColumnHandler;
import org.zkoss.zss.ui.sys.ua.impl.DeleteRowHandler;
import org.zkoss.zss.ui.sys.ua.impl.DeleteSheetHandler;
import org.zkoss.zss.ui.sys.ua.impl.FillColorHandler;
import org.zkoss.zss.ui.sys.ua.impl.FontBoldHandler;
import org.zkoss.zss.ui.sys.ua.impl.FontColorHandler;
import org.zkoss.zss.ui.sys.ua.impl.FontFamilyHandler;
import org.zkoss.zss.ui.sys.ua.impl.FontItalicHandler;
import org.zkoss.zss.ui.sys.ua.impl.FontSizeHandler;
import org.zkoss.zss.ui.sys.ua.impl.FontStrikeoutHandler;
import org.zkoss.zss.ui.sys.ua.impl.FontUnderlineHandler;
import org.zkoss.zss.ui.sys.ua.impl.HideHeaderHandler;
import org.zkoss.zss.ui.sys.ua.impl.HorizontalAlignHandler;
import org.zkoss.zss.ui.sys.ua.impl.InsertCellDownHandler;
import org.zkoss.zss.ui.sys.ua.impl.InsertCellRightHandler;
import org.zkoss.zss.ui.sys.ua.impl.InsertColumnHandler;
import org.zkoss.zss.ui.sys.ua.impl.InsertRowHandler;
import org.zkoss.zss.ui.sys.ua.impl.MergeCenterHandler;
import org.zkoss.zss.ui.sys.ua.impl.MergeHandler;
import org.zkoss.zss.ui.sys.ua.impl.MoveSheetHandler;
import org.zkoss.zss.ui.sys.ua.impl.PasteHandler;
import org.zkoss.zss.ui.sys.ua.impl.RenameSheetHandler;
import org.zkoss.zss.ui.sys.ua.impl.UnmergeHandler;
import org.zkoss.zss.ui.sys.ua.impl.VerticalAlignHandler;
import org.zkoss.zss.ui.sys.ua.impl.WrapTextHandler;
import org.zkoss.zss.undo.ClearCellAction;
import org.zkoss.zss.undo.HideHeaderAction;
import org.zkoss.zss.undo.UndoableActionManager;
/**
 * The user action handler which provide default spreadsheet operation handling.
 *  
 * @author dennis
 * @since 3.0.0
 */
public class DefaultUserActionManagerCtrl implements UserActionManagerCtrl,UserActionManager {
	private static final long serialVersionUID = 1L;
	
	public enum Category{
		AUXACTION("aux"),KEYSTROKE("key"),EVENT("event");

		private final String name;
		private Category(String name) {
			this.name = name;
		}
		
		public String getName(){
			return name;
		}
		
		public String toString(){
			return getName();
		}
	}
	
	private static final String CLIPBOARD_KEY = "$zss.clipboard$";

	private Map<String,List<UserActionHandler>> _handlerMap = new HashMap<String,List<UserActionHandler>>();
	protected static final char SPLIT_CHAR = '/';
	
	Spreadsheet _sparedsheet;//the binded spreadhseet;
	
	public void bind(Spreadsheet sparedsheet){
		this._sparedsheet = sparedsheet;
	}
	
	
	public DefaultUserActionManagerCtrl(){
		initDefaultAuxHandlers();
	}
	
	private void initDefaultAuxHandlers() {
		String category =  Category.AUXACTION.getName();
		
		//book
		registerHandler(category, AuxAction.CLOSE_BOOK.getAction(), new CloseBookHandler());
		
		//sheet
		registerHandler(category, AuxAction.ADD_SHEET.getAction(), new AddSheetHandler());
		registerHandler(category, AuxAction.DELETE_SHEET.getAction(), new DeleteSheetHandler());	
		registerHandler(category, AuxAction.RENAME_SHEET.getAction(), new RenameSheetHandler());
		registerHandler(category, AuxAction.MOVE_SHEET_LEFT.getAction(), new MoveSheetHandler(true));
		registerHandler(category, AuxAction.MOVE_SHEET_RIGHT.getAction(), new MoveSheetHandler(false));

		registerHandler(category, AuxAction.PROTECT_SHEET.getAction(), new ProtectSheetAction());
		registerHandler(category, AuxAction.GRIDLINES.getAction(), new DisplayGridlinesAction());
		
		
		registerHandler(category, AuxAction.PASTE.getAction(), new PasteHandler());
		registerHandler(category, AuxAction.CUT.getAction(), new CutHandler());
		registerHandler(category, AuxAction.COPY.getAction(), new CopyHandler());
		
		
		registerHandler(category, AuxAction.FONT_FAMILY.getAction(), new FontFamilyHandler());
		registerHandler(category, AuxAction.FONT_SIZE.getAction(), new FontSizeHandler());
		registerHandler(category, AuxAction.FONT_BOLD.getAction(), new FontBoldHandler());
		registerHandler(category, AuxAction.FONT_ITALIC.getAction(), new FontItalicHandler());
		registerHandler(category, AuxAction.FONT_UNDERLINE.getAction(), new FontUnderlineHandler());
		registerHandler(category, AuxAction.FONT_STRIKE.getAction(), new FontStrikeoutHandler());
		
		
		registerHandler(category, AuxAction.BORDER.getAction(), new ApplyBorderHandler(ApplyBorderType.EDGE_BOTTOM,BorderType.MEDIUM));
		registerHandler(category, AuxAction.BORDER_BOTTOM.getAction(), new ApplyBorderHandler(ApplyBorderType.EDGE_BOTTOM,BorderType.MEDIUM));
		registerHandler(category, AuxAction.BORDER_TOP.getAction(), new ApplyBorderHandler(ApplyBorderType.EDGE_TOP,BorderType.MEDIUM));
		registerHandler(category, AuxAction.BORDER_LEFT.getAction(), new ApplyBorderHandler(ApplyBorderType.EDGE_LEFT,BorderType.MEDIUM));
		registerHandler(category, AuxAction.BORDER_RIGHT.getAction(), new ApplyBorderHandler(ApplyBorderType.EDGE_RIGHT,BorderType.MEDIUM));
		registerHandler(category, AuxAction.BORDER_NO.getAction(), new ApplyBorderHandler(ApplyBorderType.FULL,BorderType.NONE));
		registerHandler(category, AuxAction.BORDER_ALL.getAction(), new ApplyBorderHandler(ApplyBorderType.FULL,BorderType.MEDIUM));
		registerHandler(category, AuxAction.BORDER_OUTSIDE.getAction(), new ApplyBorderHandler(ApplyBorderType.OUTLINE,BorderType.MEDIUM));
		registerHandler(category, AuxAction.BORDER_INSIDE.getAction(), new ApplyBorderHandler(ApplyBorderType.INSIDE,BorderType.MEDIUM));
		registerHandler(category, AuxAction.BORDER_INSIDE_HORIZONTAL.getAction(), new ApplyBorderHandler(ApplyBorderType.INSIDE_HORIZONTAL,BorderType.MEDIUM));
		registerHandler(category, AuxAction.BORDER_INSIDE_VERTICAL.getAction(), new ApplyBorderHandler(ApplyBorderType.INSIDE_VERTICAL,BorderType.MEDIUM));
		
		
		registerHandler(category, AuxAction.FONT_COLOR.getAction(), new FontColorHandler());
		registerHandler(category, AuxAction.FILL_COLOR.getAction(), new FillColorHandler());
		
		
		registerHandler(category, AuxAction.VERTICAL_ALIGN_TOP.getAction(), new VerticalAlignHandler(VerticalAlignment.TOP));
		registerHandler(category, AuxAction.VERTICAL_ALIGN_MIDDLE.getAction(), new VerticalAlignHandler(VerticalAlignment.CENTER));
		registerHandler(category, AuxAction.VERTICAL_ALIGN_BOTTOM.getAction(), new VerticalAlignHandler(VerticalAlignment.BOTTOM));
		registerHandler(category, AuxAction.HORIZONTAL_ALIGN_LEFT.getAction(), new HorizontalAlignHandler(Alignment.LEFT));
		registerHandler(category, AuxAction.HORIZONTAL_ALIGN_CENTER.getAction(), new HorizontalAlignHandler(Alignment.CENTER));
		registerHandler(category, AuxAction.HORIZONTAL_ALIGN_RIGHT.getAction(), new HorizontalAlignHandler(Alignment.RIGHT));
		
		registerHandler(category, AuxAction.WRAP_TEXT.getAction(), new WrapTextHandler());
		
		registerHandler(category, AuxAction.MERGE_AND_CENTER.getAction(), new MergeCenterHandler());
		registerHandler(category, AuxAction.MERGE_ACROSS.getAction(), new MergeHandler(true));
		registerHandler(category, AuxAction.MERGE_CELL.getAction(), new MergeHandler(false));
		registerHandler(category, AuxAction.UNMERGE_CELL.getAction(), new UnmergeHandler());
		
		registerHandler(category, AuxAction.INSERT_SHIFT_CELL_RIGHT.getAction(), new InsertCellRightHandler());
		registerHandler(category, AuxAction.INSERT_SHIFT_CELL_DOWN.getAction(), new InsertCellDownHandler());
		registerHandler(category, AuxAction.INSERT_SHEET_ROW.getAction(), new InsertRowHandler());
		registerHandler(category, AuxAction.INSERT_SHEET_COLUMN.getAction(), new InsertColumnHandler());
		
		registerHandler(category, AuxAction.DELETE_SHIFT_CELL_LEFT.getAction(), new DeleteCellLeftHandler());
		registerHandler(category, AuxAction.DELETE_SHIFT_CELL_UP.getAction(), new DeleteCellUpHandler());
		registerHandler(category, AuxAction.DELETE_SHEET_ROW.getAction(), new DeleteRowHandler());
		registerHandler(category, AuxAction.DELETE_SHEET_COLUMN.getAction(), new DeleteColumnHandler());
		
		registerHandler(category, AuxAction.SORT_ASCENDING.getAction(), new SortHandler(false));
		registerHandler(category, AuxAction.SORT_DESCENDING.getAction(), new SortHandler(true));
		
		registerHandler(category, AuxAction.CLEAR_CONTENT.getAction(), new ClearCellHandler(ClearCellAction.Type.CONTENT));
		registerHandler(category, AuxAction.CLEAR_STYLE.getAction(), new ClearCellHandler(ClearCellAction.Type.STYLE));
		registerHandler(category, AuxAction.CLEAR_ALL.getAction(), new ClearCellHandler(ClearCellAction.Type.ALL));
		
		registerHandler(category, AuxAction.HIDE_COLUMN.getAction(), new HideHeaderHandler(HideHeaderAction.Type.COLUMN,true));
		registerHandler(category, AuxAction.UNHIDE_COLUMN.getAction(), new HideHeaderHandler(HideHeaderAction.Type.COLUMN,false));
		registerHandler(category, AuxAction.HIDE_ROW.getAction(), new HideHeaderHandler(HideHeaderAction.Type.ROW,true));
		registerHandler(category, AuxAction.UNHIDE_ROW.getAction(), new HideHeaderHandler(HideHeaderAction.Type.ROW,false));
		
		
		
		//for enable some menu folder, do nothing
		UserActionHandler folderhandler = new AbstractProtectedHandler() {
			@Override
			protected boolean processAction(UserActionContext ctx) {
				return false;
			}
		};
		registerHandler(category, AuxAction.VERTICAL_ALIGN.getAction(), folderhandler);
		registerHandler(category, AuxAction.HORIZONTAL_ALIGN.getAction(), folderhandler);
		registerHandler(category, AuxAction.INSERT.getAction(), folderhandler);
		registerHandler(category, AuxAction.DELETE.getAction(), folderhandler);
		registerHandler(category, AuxAction.SORT_AND_FILTER.getAction(), folderhandler);
		registerHandler(category, AuxAction.CLEAR.getAction(), folderhandler);
		
		
		//key
		category =  Category.KEYSTROKE.getName();
		registerHandler(category, "^Z", new AbstractBookHandler() {
			@Override
			protected boolean processAction(UserActionContext ctx) {
				doUndo();
				return true;
			}
		});
		registerHandler(category, "^Y", new AbstractBookHandler() {
			@Override
			protected boolean processAction(UserActionContext ctx) {
				doRedo();
				return true;
			}
		});
		
		registerHandler(category, "^X", new CutHandler());
		registerHandler(category, "^C", new CopyHandler());
		registerHandler(category, "^V", new PasteHandler());
		
		
		registerHandler(category, "^B", new FontBoldHandler());
		registerHandler(category, "^I", new FontItalicHandler());
		registerHandler(category, "#del", new ClearCellHandler(ClearCellAction.Type.CONTENT));
		
		
		//event
		category =  Category.EVENT.getName();

		
	}


	protected boolean dispatchAuxAction(UserActionContext ctx) {
		boolean r = false;
		for(UserActionHandler uac :getHandlerList(ctx.getCategory(), ctx.getAction())){
			if(uac!=null && uac.isEnabled(ctx.getBook(), ctx.getSheet())){
				r |= uac.process(ctx);
			}
		}
		return r;
	}

	@Override
	public String getCtrlKeys() {
		return "^Z^Y^X^C^V^B^I^U#del";
	}
	
	protected String getAction(org.zkoss.zk.ui.event.KeyEvent event){
		StringBuilder sb = new StringBuilder();
		int keyCode = event.getKeyCode();
		boolean ctrlKey = event.isCtrlKey();
		boolean shiftKey = event.isShiftKey();
		boolean altKey = event.isAltKey();
		
		switch(keyCode){
		case KeyEvent.DELETE:
			sb.append("#del");
			break;
		}
		if(sb.length()==0 && keyCode >= 'A' && keyCode <= 'Z'){
			if(ctrlKey){
				sb.append("^");
			}
			if(altKey){
				sb.append("@");
			}
			if(shiftKey){
				sb.append("$");
			}
			sb.append(Character.toString((char)keyCode).toUpperCase());
		}
		return sb.toString();
	}
	
	protected boolean dispatchKeyAction(UserActionContext ctx) {
		KeyEvent event = (KeyEvent)ctx.getEvent();
		String action = ctx.getAction();
		if(event!=null){
			action = getAction(event);
			if(Strings.isBlank(action))
				return false;
			((UserActionContextImpl)ctx).setAction(action);
		}
		
		
		boolean r = false;
		for(UserActionHandler uac :getHandlerList(ctx.getCategory(), ctx.getAction())){
			if(uac!=null && uac.isEnabled(ctx.getBook(), ctx.getSheet())){
				r |= uac.process(ctx);
			}
		}
		return r;
	}


	//aux
	@Override
	public Set<String> getSupportedUserAction(Sheet sheet) {
		Set<String> actions = new LinkedHashSet<String>();
		
		Book book = sheet == null? null:sheet.getBook();
		String auxkey = Category.AUXACTION.getName()+SPLIT_CHAR;
		for(Entry<String, List<UserActionHandler>> entry:_handlerMap.entrySet()){
			String key = entry.getKey();
			if(key.startsWith(auxkey)){
				for(UserActionHandler handler : entry.getValue()){
					if(handler.isEnabled(book, sheet)){
						actions.add(key.substring(auxkey.length()));
						break;
					}
				}
			}
		}
		
		return actions;
	}

	@Override
	public Set<String> getInterestedEvents() {
		Set<String> evts = new LinkedHashSet<String>();
		evts.add(Events.ON_AUX_ACTION);
		evts.add(Events.ON_SHEET_SELECT);
		evts.add(Events.ON_CTRL_KEY);
		evts.add(Events.ON_CELL_SELECTION_UPDATE);
		evts.add(org.zkoss.zk.ui.event.Events.ON_CANCEL);
//		evts.add(Events.ON_CELL_DOUBLE_CLICK);
		evts.add(Events.ON_START_EDITING);
		return evts;
	}
	
	@Override
	public void onEvent(Event event) throws Exception {
		Component comp = event.getTarget();
		if(!(comp instanceof Spreadsheet)) return;
		
		Spreadsheet spreadsheet = (Spreadsheet)comp;
		Book book = spreadsheet.getBook();
		Sheet sheet = spreadsheet.getSelectedSheet();
		String action = "";
		Rect selection;
		Map extraData = null;
		if(event instanceof KeyEvent){
			//respect zss key-even't selection 
			//(could consider to remove this extra spec some day)
			selection = ((KeyEvent)event).getSelection();
//			action = "keyStroke";
		}else if(event instanceof AuxActionEvent){
			AuxActionEvent evt = ((AuxActionEvent)event);
			selection = evt.getSelection();
			//use the sheet form aux, it could be not the selected sheet
			sheet = evt.getSheet();
			action = evt.getAction();
			extraData = new HashMap(evt.getExtraData());
		}else if(event instanceof CellSelectionUpdateEvent){
			CellSelectionUpdateEvent evt = ((CellSelectionUpdateEvent)event);
			selection = new Rect(evt.getColumn(),evt.getRow(),evt.getLastColumn(),evt.getLastRow());
		}else{
			selection = spreadsheet.getSelection();
		}
		
		if(extraData==null){
			extraData = new HashMap();
		}
		
		Rect visibleSelection = new Rect(selection.getLeft(), selection.getTop(), Math.min(
			    spreadsheet.getMaxVisibleColumns(), selection.getRight()), Math.min(
			    spreadsheet.getMaxVisibleRows(), selection.getBottom()));
		
//		setLastActionData(sheet,visibleSelection,extraData);
		
		
		String nm = event.getName();
		if(event instanceof AuxActionEvent){
			UserActionContextImpl ctx = new UserActionContextImpl(_sparedsheet,event,book,sheet,visibleSelection,extraData,Category.AUXACTION.getName(),action);
			dispatchAuxAction(ctx);//aux action
		}else if(Events.ON_SHEET_SELECT.equals(nm)){
			
			updateClipboardEffect(sheet);
			//TODO 20130513, Dennis, looks like I don't need to do this here?
			//syncAutoFilter();
			
			//TODO this should be spreadsheet's job
			//toggleActionOnSheetSelected() 
		}else if(Events.ON_CTRL_KEY.equals(nm)){
			KeyEvent kevt = (KeyEvent)event;
			
			UserActionContextImpl ctx = new UserActionContextImpl(_sparedsheet,event,book,sheet,visibleSelection,extraData,Category.KEYSTROKE.getName(),action);
			
			boolean r = dispatchKeyAction(ctx);
			if(r){
				//to disable client copy/paste feature if there is any server side copy/paste
				if(kevt.isCtrlKey() && kevt.getKeyCode()=='V'){
					//internal only
					_sparedsheet.smartUpdate("doPasteFromServer", true);
				}
			}
		/*}else if(Events.ON_CELL_DOUBLE_CLICK.equals(nm)){//TODO check if we need it still
			clearClipboard();
		*/}else if(Events.ON_START_EDITING.equals(nm)){
			clearClipboard();
		}else if(org.zkoss.zk.ui.event.Events.ON_CANCEL.equals(nm)){
			clearClipboard();
		}
	}
	protected void updateClipboardEffect(Sheet sheet) {
		//to sync the 
		Clipboard cb = getClipboard();
		if (cb != null) {
			//TODO a way to know the book is different already?
			try{
				final Sheet src = cb.getSheet();
				src.getBook();//poi throw exception is sheet is already deleted.
				if(sheet.equals(src)){
					_sparedsheet.setHighlight(cb.getSelection());
				}else{
					_sparedsheet.setHighlight(null);
				}
			}catch(Exception x){
				clearClipboard();
			}
		}
	}

	@Override
	public void doAfterLoadBook(Book book) {
		clearClipboard();
	}
	
	protected Clipboard getClipboard() {
		return (Clipboard)_sparedsheet.getAttribute(CLIPBOARD_KEY);
	}
	
	protected void clearClipboard() {
		Clipboard cp = (Clipboard)_sparedsheet.removeAttribute(CLIPBOARD_KEY);
		if(cp!=null && cp.getSheet().equals(_sparedsheet.getSelectedSheet())){
			_sparedsheet.setHighlight(null);
		}
	}

	
	private String getKey(String category, String action){
		return new StringBuilder(category).append(SPLIT_CHAR).append(action).toString();
	}

	private void registerHandler(String category, String action,
			UserActionHandler handler,boolean reset) {
		if(category.indexOf(SPLIT_CHAR)>=0){
			throw new IllegalArgumentException("category can't contain "+SPLIT_CHAR+", "+category+","+action);
		}
		String key = getKey(category,action);
		List<UserActionHandler> handlers = _handlerMap.get(key);
		if(handlers==null){
			_handlerMap.put(key, handlers=new LinkedList<UserActionHandler>());
		}else if(reset){
			handlers.clear();
		}
		if(!handlers.contains(handler)){
			handlers.add(handler);
		}
	}
	@Override
	public void registerHandler(String category, String action,
			UserActionHandler handler) {
		registerHandler(category,action,handler,false);
	}
	@Override
	public void setHandler(String category, String action,
			UserActionHandler handler) {
		registerHandler(category,action,handler,true);
	}
	
	protected List<UserActionHandler> getHandlerList(String category, String action){
		
		List<UserActionHandler> list= _handlerMap.get(getKey(category,action));
		return list==null?Collections.EMPTY_LIST:list;
	}
	
	protected boolean doUndo(){
		UndoableActionManager uam = _sparedsheet.getUndoableActionManager();
		if(uam!=null && uam.isUndoable()){
			uam.undoAction();
		}
		//do we need to clear clipboard? clear clipboard will cause undo/redo misunderstanding when doing copy past with client clipboard 
//		clearClipboard();
		return true;
	}
	
	protected boolean doRedo(){
		UndoableActionManager uam = _sparedsheet.getUndoableActionManager();
		if(uam!=null && uam.isRedoable()){
			uam.redoAction();
		}
		//do we need to clear clipboard? clear clipboard will cause undo/redo misunderstanding when doing copy past with client clipboard
//		clearClipboard();
		return true;
	}
	
	protected void clearUndoable(){
		UndoableActionManager uam = _sparedsheet.getUndoableActionManager();
		if(uam!=null){
			uam.clear();
		}
	}
	
	/**
	 * internal use only
	 * @author dennis
	 *
	 */
	static public class UserActionContextImpl implements UserActionContext{

		Spreadsheet _spreadsheet;
		Book _book;
		Sheet _sheet;
		Rect _selection;
		Map<String,Object> _data;
		String _category;
		String _action;
		Event _event;
		public UserActionContextImpl(Spreadsheet ss,Event event,Book book,Sheet sheet,Rect selection,Map<String,Object> data,
				String category,String action){
			this._spreadsheet = ss;
			this._sheet = sheet;
			this._book = book;
			this._selection = selection;
			this._data = data;
			this._category = category;
			this._action = action;
			this._event = event;
		}
		
		public Book getBook(){
			return _book;
		}
		
		public Sheet getSheet(){
			return _sheet;
		}
		
		public Event getEvent(){
			return _event;
		}
		
		
		@Override
		public Spreadsheet getSpreadsheet() {
			return _spreadsheet;
		}

		@Override
		public Rect getSelection() {
			return _selection;
		}

		@Override
		public Object getData(String key) {
			return _data==null?null:_data.get(key);
		}

		@Override
		public String getCategory() {
			return _category;
		}

		@Override
		public String getAction() {
			return _action;
		}
		
		public void setAction(String action){
			this._action = action;
		}
		
		public void setData(String key,Object value){
			if(_data==null){
				_data = new HashMap<String, Object>();
			}
			_data.put(key, value);
		}

		@Override
		public Clipboard getClipboard() {
			return (Clipboard)getSpreadsheet().getAttribute(CLIPBOARD_KEY);
		}

		@Override
		public void clearClipboard() {
			Spreadsheet ss = getSpreadsheet();
			Clipboard cp = (Clipboard)ss.removeAttribute(CLIPBOARD_KEY);
			if(cp!=null && cp.getSheet().equals(ss.getSelectedSheet())){
				getSpreadsheet().setHighlight(null);
			}
		}

		@Override
		public void setClipboard(Sheet sheet, Rect selection, boolean cutMode,Object info) {
			getSpreadsheet().setAttribute(CLIPBOARD_KEY,new ClipboardImpl(sheet, selection,cutMode, info));
			if(sheet.equals(getSpreadsheet().getSelectedSheet())){
				getSpreadsheet().setHighlight(selection);
			}
		}
	}
	
	
	/**
	 * Clipboard data object for copy/paste, internal use only
	 * @author dennis
	 *
	 */
	public static class ClipboardImpl implements Clipboard{

		final Rect _selection;
		final Sheet _sheet;
		final boolean _cutMode;
		final Object _info;
		
		public ClipboardImpl(Sheet sheet,Rect selection, boolean cutMode,Object info) {
			if(sheet==null){
				throw new IllegalArgumentException("Sheet is null");
			}
			if(selection==null){
				throw new IllegalArgumentException("selection is null");
			}
			this._sheet = sheet;
			this._selection = selection;
			this._cutMode = cutMode;
			this._info = info;
		}

		@Override
		public Sheet getSheet() {
			return _sheet;
		}

		@Override
		public Rect getSelection() {
			return _selection;
		}

		@Override
		public Object getInfo() {
			return _info;
		}
		
		public boolean isCutMode(){
			return _cutMode;
		}
	}

}
