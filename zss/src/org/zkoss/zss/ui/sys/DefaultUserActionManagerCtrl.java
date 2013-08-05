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
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
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
import org.zkoss.zss.ui.ua.AddSheetHandler;
import org.zkoss.zss.ui.ua.CloseBookHandler;
import org.zkoss.zss.ui.ua.DeleteSheetHandler;
import org.zkoss.zss.ui.ua.MoveSheetHandler;
import org.zkoss.zss.ui.ua.NilHandler;
import org.zkoss.zss.ui.ua.RenameSheetHandler;
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
	
	public enum PasteType {
		COPY,
		CUT
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
		UserActionHandler uah;
		
		//book
		registerHandler(category, AuxAction.CLOSE_BOOK.getAction(), new CloseBookHandler());
		
		//sheet
		registerHandler(category, AuxAction.ADD_SHEET.getAction(), new AddSheetHandler());
		registerHandler(category, AuxAction.DELETE_SHEET.getAction(), new DeleteSheetHandler());	
		registerHandler(category, AuxAction.RENAME_SHEET.getAction(), new RenameSheetHandler());
		registerHandler(category, AuxAction.MOVE_SHEET_LEFT.getAction(), new MoveSheetHandler(true));
		registerHandler(category, AuxAction.MOVE_SHEET_RIGHT.getAction(), new MoveSheetHandler(false));

		registerHandler(category, AuxAction.PROTECT_SHEET.getAction(), new NilHandler());
		registerHandler(category, AuxAction.GRIDLINES.getAction(), new NilHandler());
		
		
		registerHandler(category, AuxAction.PASTE.getAction(), new NilHandler());
		registerHandler(category, AuxAction.PASTE_FORMULA.getAction(), new NilHandler());
		registerHandler(category, AuxAction.PASTE_VALUE.getAction(), new NilHandler());
		registerHandler(category, AuxAction.PASTE_ALL_EXPECT_BORDERS.getAction(), new NilHandler());
		registerHandler(category, AuxAction.PASTE_TRANSPOSE.getAction(), new NilHandler());
		registerHandler(category, AuxAction.CUT.getAction(), new NilHandler());
		registerHandler(category, AuxAction.COPY.getAction(), new NilHandler());
		
		
		registerHandler(category, AuxAction.FONT_FAMILY.getAction(), new NilHandler());
		registerHandler(category, AuxAction.FONT_SIZE.getAction(), new NilHandler());
		registerHandler(category, AuxAction.FONT_BOLD.getAction(), new NilHandler());
		registerHandler(category, AuxAction.FONT_ITALIC.getAction(), new NilHandler());
		registerHandler(category, AuxAction.FONT_UNDERLINE.getAction(), new NilHandler());
		registerHandler(category, AuxAction.FONT_STRIKE.getAction(), new NilHandler());
		
		
		registerHandler(category, AuxAction.BORDER.getAction(), new NilHandler());
		registerHandler(category, AuxAction.BORDER_BOTTOM.getAction(), new NilHandler());
		registerHandler(category, AuxAction.BORDER_TOP.getAction(), new NilHandler());
		registerHandler(category, AuxAction.BORDER_LEFT.getAction(), new NilHandler());
		registerHandler(category, AuxAction.BORDER_RIGHT.getAction(), new NilHandler());
		registerHandler(category, AuxAction.BORDER_NO.getAction(), new NilHandler());
		registerHandler(category, AuxAction.BORDER_ALL.getAction(), new NilHandler());
		registerHandler(category, AuxAction.BORDER_OUTSIDE.getAction(), new NilHandler());
		registerHandler(category, AuxAction.BORDER_INSIDE.getAction(), new NilHandler());
		registerHandler(category, AuxAction.BORDER_INSIDE_HORIZONTAL.getAction(), new NilHandler());
		registerHandler(category, AuxAction.BORDER_INSIDE_VERTICAL.getAction(), new NilHandler());
		
		
		registerHandler(category, AuxAction.FONT_COLOR.getAction(), new NilHandler());
		registerHandler(category, AuxAction.FILL_COLOR.getAction(), new NilHandler());
		
		
		registerHandler(category, AuxAction.VERTICAL_ALIGN_TOP.getAction(), new NilHandler());
		registerHandler(category, AuxAction.VERTICAL_ALIGN_MIDDLE.getAction(), new NilHandler());
		registerHandler(category, AuxAction.VERTICAL_ALIGN_BOTTOM.getAction(), new NilHandler());
		registerHandler(category, AuxAction.HORIZONTAL_ALIGN_LEFT.getAction(), new NilHandler());
		registerHandler(category, AuxAction.HORIZONTAL_ALIGN_CENTER.getAction(), new NilHandler());
		registerHandler(category, AuxAction.HORIZONTAL_ALIGN_RIGHT.getAction(), new NilHandler());
		
		registerHandler(category, AuxAction.WRAP_TEXT.getAction(), new NilHandler());
		
		registerHandler(category, AuxAction.MERGE_AND_CENTER.getAction(), new NilHandler());
		registerHandler(category, AuxAction.MERGE_ACROSS.getAction(), new NilHandler());
		registerHandler(category, AuxAction.MERGE_CELL.getAction(), new NilHandler());
		registerHandler(category, AuxAction.UNMERGE_CELL.getAction(), new NilHandler());
		
		registerHandler(category, AuxAction.INSERT_SHIFT_CELL_RIGHT.getAction(), new NilHandler());
		registerHandler(category, AuxAction.INSERT_SHIFT_CELL_DOWN.getAction(), new NilHandler());
		registerHandler(category, AuxAction.INSERT_SHEET_ROW.getAction(), new NilHandler());
		registerHandler(category, AuxAction.INSERT_SHEET_COLUMN.getAction(), new NilHandler());
		
		registerHandler(category, AuxAction.DELETE_SHIFT_CELL_LEFT.getAction(), new NilHandler());
		registerHandler(category, AuxAction.DELETE_SHIFT_CELL_UP.getAction(), new NilHandler());
		registerHandler(category, AuxAction.DELETE_SHEET_ROW.getAction(), new NilHandler());
		registerHandler(category, AuxAction.DELETE_SHEET_COLUMN.getAction(), new NilHandler());
		
		registerHandler(category, AuxAction.SORT_ASCENDING.getAction(), new NilHandler());
		registerHandler(category, AuxAction.SORT_DESCENDING.getAction(), new NilHandler());
		
		registerHandler(category, AuxAction.FILTER.getAction(), new NilHandler());
		registerHandler(category, AuxAction.CLEAR_FILTER.getAction(), new NilHandler());
		registerHandler(category, AuxAction.REAPPLY_FILTER.getAction(), new NilHandler());
		
		registerHandler(category, AuxAction.CLEAR_CONTENT.getAction(), new NilHandler());
		registerHandler(category, AuxAction.CLEAR_STYLE.getAction(), new NilHandler());
		registerHandler(category, AuxAction.CLEAR_ALL.getAction(), new NilHandler());
		
		registerHandler(category, AuxAction.HIDE_COLUMN.getAction(), new NilHandler());
		registerHandler(category, AuxAction.UNHIDE_COLUMN.getAction(), new NilHandler());
		registerHandler(category, AuxAction.HIDE_ROW.getAction(), new NilHandler());
		registerHandler(category, AuxAction.UNHIDE_ROW.getAction(), new NilHandler());
		
		
		
		//key
		category =  Category.KEYSTROKE.getName();
		registerHandler(category, "^Z", new NilHandler());
		registerHandler(category, "^Y", new NilHandler());
		registerHandler(category, "^X", new NilHandler());
		registerHandler(category, "^C", new NilHandler());
		registerHandler(category, "^V", new NilHandler());
		registerHandler(category, "^D", new NilHandler());
		registerHandler(category, "^B", new NilHandler());
		registerHandler(category, "^I", new NilHandler());
		registerHandler(category, "#del", new NilHandler());
		
		
		//event
		category =  Category.EVENT.getName();
		registerHandler(category, Events.ON_CELL_SELECTION_UPDATE+SPLIT_CHAR+"move", new NilHandler());
		registerHandler(category, Events.ON_CELL_SELECTION_UPDATE+SPLIT_CHAR+"resize", new NilHandler());

		
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
		return "^Z^Y^X^C^V^D^B^I^U#del";
	}
	
	protected String getAction(KeyEvent event){
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
	
	
	protected boolean dispatchCellSelectionUpdateAction(UserActionContext ctx) {
		CellSelectionUpdateEvent evt = (CellSelectionUpdateEvent)ctx.getEvent();
		//last selection either get form selection or from event
		String action;
		if(evt.getAction()==CellSelectionAction.MOVE){
			action = Events.ON_CELL_SELECTION_UPDATE+SPLIT_CHAR+"move";
		}else if(evt.getAction()==CellSelectionAction.RESIZE){
			action = Events.ON_CELL_SELECTION_UPDATE+SPLIT_CHAR+"resize";
		}else{
			return false;
		}
		
//		doResizeCellSelection(new Rect(evt.getOrigleft(),evt.getOrigtop(),evt.getOrigright(),evt.getOrigbottom()),ctx.getSelection());
//		doMoveCellSelection(new Rect(evt.getOrigleft(),evt.getOrigtop(),evt.getOrigright(),evt.getOrigbottom()),ctx.getSelection());
		
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
		}else if(Events.ON_CELL_SELECTION_UPDATE.equals(nm)){
			
			UserActionContextImpl ctx = new UserActionContextImpl(_sparedsheet,event,book,sheet,visibleSelection,extraData,Category.EVENT.getName(),action);
			
			dispatchCellSelectionUpdateAction(ctx);
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
		System.out.println(">>>> registerHandler "+category+","+action+":"+handler);
		
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
		System.out.println(">>>> getHandler "+category+","+action+":"+list);
		return list;
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
		public void setClipboard(Sheet sheet, Rect selection, Object info) {
			getSpreadsheet().setAttribute(CLIPBOARD_KEY,new ClipboardImpl(sheet, selection, info));
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

		public final Rect _selection;
		public final Sheet _sheet;
		public final Object _info;
		
		public ClipboardImpl(Sheet sheet,Rect selection,Object info) {
			if(sheet==null){
				throw new IllegalArgumentException("Sheet is null");
			}
			if(selection==null){
				throw new IllegalArgumentException("selection is null");
			}
			this._sheet = sheet;
			this._selection = selection;
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
	}

}
