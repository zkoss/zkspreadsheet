package org.zkoss.zss.essential.events;

import java.util.ArrayList;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.AuxActionEvent;
import org.zkoss.zss.ui.event.CellAreaEvent;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.CellFilterEvent;
import org.zkoss.zss.ui.event.CellHyperlinkEvent;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zss.ui.event.CellSelectionUpdateEvent;
import org.zkoss.zss.ui.event.EditboxEditingEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.HeaderMouseEvent;
import org.zkoss.zss.ui.event.HeaderUpdateEvent;
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
 * This controller listens all Spreadsheet's events and show related messages.
 * @author dennis, Hawk
 *
 */
public class EventsComposer extends SelectorComposer<Component>{

	private static final long serialVersionUID = 1L;
	
	private ListModelList<String> eventFilterModel = new ListModelList<String>();
	private ListModelList<String> infoModel = new ListModelList<String>();
	
	@Wire("spreadsheet")
	Spreadsheet spreadsheet;
	@Wire
	private Listbox eventFilterList;
	@Wire
	private Grid infoList;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		initModel();
	}

	//Events

	@Listen("onStartEditing = spreadsheet")
	public void onStartEditing(StartEditingEvent event){
		StringBuilder info = new StringBuilder();
		String ref = Ranges.getCellReference(event.getRow(),event.getColumn());
		info.append("Start editing ").append(ref)
		.append(", editing-value is ").append("\""+event.getEditingValue()+"\"")
		.append(" client-value is ").append("\""+event.getClientValue()+"\"");
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
		
		if(ref.equals("D1")){
			String newValue = "Surprise!!";
			//we change the editing value
			event.setEditingValue(newValue);
			addInfo("Editing value is change to "+newValue);
		}else if(ref.equals("E1")){
			//forbid editing
			event.cancel();
			addInfo("Editing E1 is canceled");
		}
	}

	@Listen("onEditboxEditing = spreadsheet")
	public void onEditboxEditing(EditboxEditingEvent event){
		StringBuilder info = new StringBuilder();
		String ref = Ranges.getCellReference(event.getRow(),event.getColumn());
		info.append("Editing ").append(ref)
		.append(", value is ").append("\""+event.getEditingValue()+"\"");
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}	
	
	@Listen("onStopEditing = spreadsheet")
	public void onStopEditing(StopEditingEvent event){
		StringBuilder info = new StringBuilder();
		String ref = Ranges.getCellReference(event.getRow(),event.getColumn());
		info.append("Stop editing ").append(ref)
		.append(", editing-value is ").append("\""+event.getEditingValue()+"\"");
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
		
		if(ref.equals("D3")){
			String newValue = event.getEditingValue()+"-Woo";
			//we change the editing value
			event.setEditingValue(newValue);
			addInfo("Editing value is changed to \""+newValue+"\"");
		}else if(ref.equals("E3")){
			//forbid editing
			event.cancel();
			addInfo("Editing E3 is canceled");
		}
	}
	
	@Listen("onCellChange = spreadsheet")
	public void onCellChange(CellAreaEvent event){
		StringBuilder info = new StringBuilder();

		info.append("Cell changes on ").append(Ranges.getAreaReference(event.getArea()));
		info.append(", first value is \""
		+Ranges.range(event.getSheet(),event.getArea()).getCellFormatText()+"\"");

		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCtrlKey = spreadsheet")
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
	
	@Listen("onCellClick = spreadsheet")
	public void onCellClick(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Click on cell ")
		.append(Ranges.getCellReference(event.getRow(),event.getColumn()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	@Listen("onCellRightClick = spreadsheet")
	public void onCellRightClick(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Right-click on cell ").append(Ranges.getCellReference(event.getRow(),event.getColumn()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	@Listen("onCellDoubleClick = spreadsheet")
	public void onCellDoubleClick(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Double-click on cell ").append(Ranges.getCellReference(event.getRow(),event.getColumn()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	
	
	@Listen("onHeaderClick = spreadsheet")
	public void onHeaderClick(HeaderMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Click on ").append(event.getType()).append(" ");
		
		switch(event.getType()){
		case COLUMN:
			info.append(Ranges.getColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(Ranges.getRowReference(event.getIndex()));
			break;
		}
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	@Listen("onHeaderRightClick = spreadsheet")
	public void onHeaderRightClick(HeaderMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Right-click on ").append(event.getType()).append(" ");
		
		switch(event.getType()){
		case COLUMN:
			info.append(Ranges.getColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(Ranges.getRowReference(event.getIndex()));
			break;
		}
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	@Listen("onHeaderDoubleClick = spreadsheet")
	public void onHeaderDoubleClick(HeaderMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Double-click on ").append(event.getType()).append(" ");
		
		switch(event.getType()){
		case COLUMN:
			info.append(Ranges.getColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(Ranges.getRowReference(event.getIndex()));
			break;
		}
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}

	@Listen("onHeaderUpdate = spreadsheet")
	public void onHeaderUpdate(HeaderUpdateEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Header ").append(event.getAction())
			.append(" on ").append(event.getType());
		switch(event.getType()){
		case COLUMN:
			info.append(" ").append(Ranges.getColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(" ").append(Ranges.getRowReference(event.getIndex()));
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
	
	@Listen("onCellFocused = spreadsheet")
	public void onCellFocused(CellEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Focus on[").append(Ranges.getCellReference(event.getRow(),event.getColumn())).append("]");
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCellSelection = spreadsheet")
	public void onCellSelection(CellSelectionEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Select on[").append(Ranges.getAreaReference(event.getArea())).append("]");
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCellSelectionUpdate = spreadsheet")
	public void onCellSelectionUpdate(CellSelectionUpdateEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Selection update from[")
		.append(Ranges.getAreaReference(event.getOrigRow(),event.getOrigColumn()
				, event.getOrigLastRow(),event.getOrigLastColumn()))
		.append("] to [")
		.append(Ranges.getAreaReference(event.getArea())).append("]");

		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}

	@Listen("onCellHyperlink = spreadsheet")
	public void onCellHyperlink(CellHyperlinkEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("Hyperlink ").append(event.getType())
			.append(" on : ").append(Ranges.getCellReference(event.getRow(),event.getColumn()))
			.append(", address : ").append(event.getAddress());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}		

	

	
	@Listen("onSheetSelect = spreadsheet")
	public void onSheetSelect(SheetSelectEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Select sheet : ").append(event.getSheetName());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onSheetNameChange = spreadsheet")
	public void onSheetNameChange(SheetEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Rename sheet to ").append(event.getSheetName());
		
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onSheetOrderChange = spreadsheet")
	public void onSheetOrderChange(SheetEvent event){
		StringBuilder info = new StringBuilder();
		Sheet sheet = event.getSheet();
		info.append("Reorder sheet : ").append(event.getSheetName()).append(" to ").append(sheet.getBook().getSheetIndex(sheet));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onSheetCreate = spreadsheet")
	public void onSheetCreate(SheetEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Create sheet : ").append(event.getSheetName());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onSheetDelete = spreadsheet")
	public void onSheetDelete(SheetDeleteEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Delete sheet : ").append(event.getSheetName());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	

	
	@Listen("onWidgetUpdate = spreadsheet")
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
	
	@Listen("onWidgetCtrlKey = spreadsheet")
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
	
	@Listen("onAuxAction = spreadsheet")
	public void onAuxAction(AuxActionEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("AuxAction ").append(event.getAction())
			.append(" on : ").append(Ranges.getAreaReference(event.getSheet(),event.getSelection()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	
	@Listen("onCellFilter = spreadsheet")
	public void onCellFilter(CellFilterEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("Filter button clicked")
			.append(", filter area: ").append(Ranges.getAreaReference(event.getFilterArea()))
			.append(", on field: ").append(event.getField());
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	
	@Listen("onCellValidator = spreadsheet")
	public void onCellValidator(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		
		info.append("Validation button clicked ")
			.append(" on cell ").append(Ranges.getCellReference(event.getRow(),event.getColumn()));
		
		if(isShowEventInfo(event.getName())){
			addInfo(info.toString());
		}
	}
	

	private void initModel() {
		//Available events
		//It is just for showing message, event is always listened in this demo.
		eventFilterModel.setMultiple(true);
		
		addEventFilter(Events.ON_START_EDITING,true);
		addEventFilter(Events.ON_EDITBOX_EDITING,true);
		addEventFilter(Events.ON_STOP_EDITING,true);
		addEventFilter(Events.ON_CELL_CHANGE,true);
		
		addEventFilter(Events.ON_CTRL_KEY,true);
		
		addEventFilter(Events.ON_CELL_CLICK,false);
		addEventFilter(Events.ON_CELL_DOUBLE_CLICK,true);
		addEventFilter(Events.ON_CELL_RIGHT_CLICK,true);
		
		addEventFilter(Events.ON_HEADER_UPDATE,true);
		addEventFilter(Events.ON_HEADER_CLICK,true);
		addEventFilter(Events.ON_HEADER_RIGHT_CLICK,true);
		addEventFilter(Events.ON_HEADER_DOUBLE_CLICK,true);
		
		addEventFilter(Events.ON_AUX_ACTION,true);
		
		addEventFilter(Events.ON_CELL_FOUCSED,false);
		addEventFilter(Events.ON_CELL_SELECTION,false);
		addEventFilter(Events.ON_CELL_SELECTION_UPDATE,true);
		
		addEventFilter(Events.ON_CELL_FILTER,true);//useless
		addEventFilter(Events.ON_CELL_VALIDATOR,true);//useless
		
		addEventFilter(Events.ON_WIDGET_UPDATE,true);
		addEventFilter(Events.ON_WIDGET_CTRL_KEY,true);
		
		addEventFilter(Events.ON_SHEET_CREATE,true);
		addEventFilter(Events.ON_SHEET_DELETE,true);
		addEventFilter(Events.ON_SHEET_NAME_CHANGE,true);
		addEventFilter(Events.ON_SHEET_ORDER_CHANGE,true);
		addEventFilter(Events.ON_SHEET_SELECT,true);
		
		addEventFilter(Events.ON_CELL_HYPERLINK,true);	
		
		eventFilterList.setModel(eventFilterModel);

		//add default show only
		infoList.setModel(infoModel);
		addInfo("Spreadsheet initialized");		
		
		//hint for special cells
		//D1
		Ranges.range(spreadsheet.getSelectedSheet(), 0, 3).setCellEditText("Edit Me");
		//E1
		Ranges.range(spreadsheet.getSelectedSheet(), 0, 4).setCellEditText("Edit Me");
		//D3
		Ranges.range(spreadsheet.getSelectedSheet(), 2, 3).setCellEditText("Edit Me");
		//E3
		Ranges.range(spreadsheet.getSelectedSheet(), 2, 4).setCellEditText("Edit Me");
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

	@Listen("onClick = #clearAllFilter")
	public void onClearAllFilter(){
		eventFilterModel.clearSelection();
	}
	@Listen("onClick = #selectAllFilter")
	public void onSelectAll(){
		eventFilterModel.clearSelection();
		eventFilterModel.setSelection(new ArrayList<String>(eventFilterModel));
	}
	
	private void addInfo(String info){
		infoModel.add(0, info);
		while(infoModel.size()>100){
			infoModel.remove(infoModel.size()-1);
		}
	}

	@Listen("onClick = #clearInfo")
	public void onClearInfo(){
		infoModel.clear();
	}
	/**
	 * ON_CELL_FILTER //useless
	 * ON_CELL_VALIDATOR //useless
	 */
}



