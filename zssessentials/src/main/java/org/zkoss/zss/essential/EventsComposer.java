package org.zkoss.zss.essential;

import java.io.IOException;
import java.util.ArrayList;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.DefaultUserAction;
import org.zkoss.zss.ui.event.AuxActionEvent;
import org.zkoss.zss.ui.event.CellAreaEvent;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.CellFilterEvent;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zss.ui.event.CellSelectionUpdateEvent;
import org.zkoss.zss.ui.event.EditboxEditingEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.HeaderMouseEvent;
import org.zkoss.zss.ui.event.HeaderUpdateEvent;
import org.zkoss.zss.ui.event.CellHyperlinkEvent;
import org.zkoss.zss.ui.event.KeyEvent;
import org.zkoss.zss.ui.event.SheetDeleteEvent;
import org.zkoss.zss.ui.event.SheetEvent;
import org.zkoss.zss.ui.event.SheetSelectEvent;
import org.zkoss.zss.ui.event.StartEditingEvent;
import org.zkoss.zss.ui.event.StopEditingEvent;
import org.zkoss.zss.ui.event.WidgetKeyEvent;
import org.zkoss.zss.ui.event.WidgetUpdateEvent;
import org.zkoss.zul.Grid;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class EventsComposer extends SelectorComposer<Component> {

	private static final long serialVersionUID = 1L;

	ListModelList<String> infoModel = new ListModelList<String>();
	
	ListModelList<String> availableBookModel = new ListModelList<String>();
	
	ListModelList<String> eventFilterModel = new ListModelList<String>();
	
	@Wire
	Grid infoList;
	
	@Wire
	Listbox availableBookList;
	
	@Wire
	Listbox eventFilterList;
	
	@Wire
	Spreadsheet ss;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		initModel();
		
		availableBookList.setModel(availableBookModel);
		infoList.setModel(infoModel);
		eventFilterList.setModel(eventFilterModel);
		
		addInfo("Spreadsheet initialized");
		
		String book = Executions.getCurrent().getParameter("book");
		if(book!=null){
			try{
				loadBookFromAvailable(book);
			}catch(Exception x){}
		}
		
		String sheet = Executions.getCurrent().getParameter("sheet");
		if(sheet!=null){
			try{
				Sheet ssheet = ss.getBook().getSheet(sheet);
				if(ssheet!=null){
					ss.setSelectedSheet(ssheet.getSheetName());
				}
			}catch(Exception x){}
		} 
		
		//sync available book selection
		book = ss.getBook().getBookName();
		
		availableBookModel.addToSelection(book);
	}

	private void initModel() {
		availableBookModel.add("blank.xlsx");
		availableBookModel.add("sample.xlsx");
		
		
		//Available events
		//It is just for showing message, event is always listened in this demo.
		eventFilterModel.setMultiple(true);
		
		addEventFilter(Events.ON_AUX_ACTION,true);
		
		addEventFilter(Events.ON_CELL_CHANGE,true);
		addEventFilter(Events.ON_CELL_FOUCSED,false);
		addEventFilter(Events.ON_CELL_SELECTION,false);
		addEventFilter(Events.ON_CELL_SELECTION_UPDATE,true);
		
		
		
		addEventFilter(Events.ON_CELL_CLICK,false);
		addEventFilter(Events.ON_CELL_DOUBLE_CLICK,true);
		addEventFilter(Events.ON_CELL_RIGHT_CLICK,true);
		
		addEventFilter(Events.ON_CELL_FILTER,true);//useless
		addEventFilter(Events.ON_CELL_VALIDATOR,true);//useless
		
		addEventFilter(Events.ON_START_EDITING,true);
		addEventFilter(Events.ON_EDITBOX_EDITING,true);
		addEventFilter(Events.ON_STOP_EDITING,true);
		
		
		addEventFilter(Events.ON_HEADER_UPDATE,true);
		addEventFilter(Events.ON_HEADER_CLICK,true);
		addEventFilter(Events.ON_HEADER_RIGHT_CLICK,true);
		addEventFilter(Events.ON_HEADER_DOUBLE_CLICK,true);
		
		addEventFilter(Events.ON_CTRL_KEY,true);
		
		addEventFilter(Events.ON_WIDGET_UPDATE,true);
		addEventFilter(Events.ON_WIDGET_CTRL_KEY,true);
		
		addEventFilter(Events.ON_SHEET_CREATE,true);
		addEventFilter(Events.ON_SHEET_DELETE,true);
		addEventFilter(Events.ON_SHEET_NAME_CHANGE,true);
		addEventFilter(Events.ON_SHEET_ORDER_CHANGE,true);
		addEventFilter(Events.ON_SHEET_SELECT,true);
		
		addEventFilter(Events.ON_CELL_HYPERLINK,true);	
		
		//add default show only
		
	}
	
	private void addEventFilter(String event,boolean showinfo){
		if(!eventFilterModel.contains(eventFilterModel)){
			eventFilterModel.add(event);
		}
		if(showinfo){
			eventFilterModel.addToSelection(event);
		}else{
			eventFilterModel.removeFromSelection(event);
		}
	}
	
	private boolean isShowEventInfo(String event){
		return eventFilterModel.getSelection().contains(event);
	}

	private void addInfo(String info){
		infoModel.add(0, info);
		while(infoModel.size()>100){
			infoModel.remove(infoModel.size()-1);
		}
	}
	
	@Listen("onSelect = #availableBookList")
	public void onBookSelect(){
		String bookname = availableBookModel.getSelection().iterator().next();
		loadBookFromAvailable(bookname);
	}
	
	protected void loadBookFromAvailable(int index){
		String bookname = availableBookModel.get(index);
		loadBookFromAvailable(bookname);
	}
	
	protected void loadBookFromAvailable(String bookname){
		if(!availableBookModel.contains(bookname)){
			return;
		}
		Importer imp = Importers.getImporter();
		try {
			Book book = imp.imports(WebApps.getCurrent().getResource("/WEB-INF/books/"+bookname), bookname);
			ss.setBook(book);
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage(),e);
		}
	}	
	
	
	@Listen("onClick = #clearInfo")
	public void onClearInfo(){
		infoModel.clear();
	}
	@Listen("onClick = #clearAllFilter")
	public void onClearAllFilter(){
		eventFilterModel.clearSelection();
	}
	@Listen("onClick = #selectAllFilter")
	public void onSelectAll(){
		eventFilterModel.clearSelection();
		eventFilterModel.setSelection(new ArrayList<String>(eventFilterModel));
	}
	@Listen("onAuxAction = #ss")
	public void onAuxActionHandling(AuxActionEvent event){
		//handle extra action when book close
		if(event.getAction().equals(DefaultUserAction.CLOSE_BOOK.toString())){
			availableBookModel.clearSelection();
		}
	}
	
	
	//Events
	
	@Listen("onCellFocused = #ss")
	public void onCellFocused(CellEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Focus on[").append(Ranges.getCellReference(event.getRow(),event.getColumn())).append("]");
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCellSelection = #ss")
	public void onCellSelection(CellSelectionEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Select on[").append(Ranges.getAreaReference(event.getArea())).append("]");
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCellSelectionUpdate = #ss")
	public void onCellSelectionUpdate(CellSelectionUpdateEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Selection update from[")
				.append(Ranges.getAreaReference(event.getOrigRow(),event.getOrigColumn(), event.getOrigLastRow(),event.getOrigLastColumn())).append("] to [")
				.append(Ranges.getAreaReference(event.getArea())).append("]");
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}

	@Listen("onHeaderUpdate = #ss")
	public void onHeaderUpdate(HeaderUpdateEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Header ").append(event.getAction())
			.append(" on ").append(event.getType());
		switch(event.getType()){
		case COLUMN:
			info.append(" ").append(Ranges.getColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(" ").append(Ranges.toRowReference(event.getIndex()));
			break;
		}

		switch(event.getAction()){
		case RESIZE:
			if(event.isHidden()){
				info.append(" hides ");
			}else{
				info.append(" changes to ").append(event.getSize());
			}
			break;
		}
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	
	@Listen("onHeaderClick = #ss")
	public void onHeaderClick(HeaderMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Click on ").append(event.getType()).append(" ");
		
		switch(event.getType()){
		case COLUMN:
			info.append(Ranges.getColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(Ranges.toRowReference(event.getIndex()));
			break;
		}
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	@Listen("onHeaderRightClick = #ss")
	public void onHeaderRightClick(HeaderMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Right-click on ").append(event.getType()).append(" ");
		
		switch(event.getType()){
		case COLUMN:
			info.append(Ranges.getColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(Ranges.toRowReference(event.getIndex()));
			break;
		}
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	@Listen("onHeaderDoubleClick = #ss")
	public void onHeaderDoubleClick(HeaderMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Double-click on ").append(event.getType()).append(" ");
		
		switch(event.getType()){
		case COLUMN:
			info.append(Ranges.getColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(Ranges.toRowReference(event.getIndex()));
			break;
		}
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCellClick = #ss")
	public void onCellClick(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Click on cell ").append(Ranges.getCellReference(event.getRow(),event.getColumn()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	@Listen("onCellRightClick = #ss")
	public void onCellRightClick(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Right-click on cell ").append(Ranges.getCellReference(event.getRow(),event.getColumn()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	@Listen("onCellDoubleClick = #ss")
	public void onCellDoubleClick(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Double-click on cell ").append(Ranges.getCellReference(event.getRow(),event.getColumn()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onStartEditing = #ss")
	public void onStartEditing(StartEditingEvent event){
		StringBuilder info = new StringBuilder();
		String ref = Ranges.getCellReference(event.getRow(),event.getColumn());
		info.append("Start to edit ").append(ref)
		.append(", editing-value is ").append(event.getEditingValue()).append(" client-value is ").append(event.getClientValue());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
		
		if(ref.equals("D1")){
			String newval = "Supprise!!";
			//we change the editing value
			event.setEditingValue(newval);
			addInfo("Editing value is change to "+newval);
		}else if(ref.equals("E1")){
			//we don't allow edit D2
			event.cancel();
			addInfo("Editing is canceled");
		}else{
			addInfo("Try to edit D1 or E1");
		}
	}
	
	@Listen("onStopEditing = #ss")
	public void onStopEditing(StopEditingEvent event){
		StringBuilder info = new StringBuilder();
		String ref = Ranges.getCellReference(event.getRow(),event.getColumn());
		info.append("Stop editing ").append(ref)
		.append(", value is ").append(event.getEditingValue());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
		
		if(ref.equals("D3")){
			String newval = event.getEditingValue()+"-Woo";
			//we change the editing value
			event.setEditingValue(newval);
			addInfo("Editing value is changing to "+newval);
		}else if(ref.equals("E3")){
			//we don't allow edit D2
			event.cancel();
			addInfo("Editing is canceled");
		}else{
			addInfo("Try to edit D3 or E3");
		}
	}
	
	@Listen("onEditboxEditing = #ss")
	public void onEditboxEditing(EditboxEditingEvent event){
		StringBuilder info = new StringBuilder();
		String ref = Ranges.getCellReference(event.getRow(),event.getColumn());
		info.append("Editing ").append(ref)
		.append(", value is ").append(event.getEditingValue());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onSheetSelect = #ss")
	public void onSheetSelect(SheetSelectEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Select sheet : ").append(event.getSheetName());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onSheetNameChange = #ss")
	public void onSheetNameChange(SheetEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Rename sheet to ").append(event.getSheetName());
		
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onSheetOrderChange = #ss")
	public void onSheetOrderChange(SheetEvent event){
		StringBuilder info = new StringBuilder();
		Sheet sheet = event.getSheet();
		info.append("Reorder sheet : ").append(event.getSheetName()).append(" to ").append(sheet.getBook().getSheetIndex(sheet));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onSheetCreate = #ss")
	public void onSheetCreate(SheetEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Create sheet : ").append(event.getSheetName());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onSheetDelete = #ss")
	public void onSheetDelete(SheetDeleteEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Delete sheet : ").append(event.getSheetName());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCellChange = #ss")
	public void onCellChange(CellAreaEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("Cell changes on ").append(Ranges.getAreaReference(event.getArea()));
		info.append(", first value is :"+Ranges.range(event.getSheet(),event.getArea()).getCellFormatText());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCtrlKey = #ss")
	public void onCtrlKey(KeyEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("Keys : ").append(event.getKeyCode())
			.append(", Ctrl:").append(event.isCtrlKey())
			.append(", Alt:").append(event.isAltKey())
			.append(", Shift:").append(event.isShiftKey());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onWidgetUpdate = #ss")
	public void onWidgetUpdate(WidgetUpdateEvent event){
		StringBuilder info = new StringBuilder();
		SheetAnchor anchor = event.getSheetAnchor();
		
		info.append("Widget ")
				.append(event.getWidgetData())
				.append(" ")
				.append(event.getAction())
				.append(" to ")
				.append(Ranges.getAreaReference(anchor.getRow(),
						anchor.getColumn(), anchor.getLastRow(),
						anchor.getLastColumn()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onWidgetCtrlKey = #ss")
	public void onWidgetCtrlKey(WidgetKeyEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("Widget ").append(event.getWidgetData())
			.append(" Key : ").append(event.getKeyCode())
			.append(", Ctrl:").append(event.isCtrlKey())
			.append(", Alt:").append(event.isAltKey())
			.append(", Shift:").append(event.isShiftKey());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onAuxAction = #ss")
	public void onAuxAction(AuxActionEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("AuxAction ").append(event.getAction())
			.append(" on : ").append(Ranges.getAreaReference(event.getSheet(),event.getSelection()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCellHyperlink = #ss")
	public void onCellHyperlink(CellHyperlinkEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("Hyperlink ").append(event.getType())
			.append(" on : ").append(Ranges.getCellReference(event.getRow(),event.getColumn()))
			.append(", address : ").append(event.getAddress());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}	
	
	@Listen("onCellFilter = #ss")
	public void onCellFilter(CellFilterEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("Filter button clicked")
			.append(", filter area: ").append(Ranges.getAreaReference(event.getFilterArea()))
			.append(", on field: ").append(event.getField());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCellValidator = #ss")
	public void onCellValidator(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("Validation button clicked ")
			.append(" on cell ").append(Ranges.getCellReference(event.getRow(),event.getColumn()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	
	/**
	 * ON_CELL_FILTER //useless
	 * ON_CELL_VALIDATOR //useless
	 */
}



