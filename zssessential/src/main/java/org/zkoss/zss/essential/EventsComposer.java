package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zss.ui.event.CellSelectionUpdateEvent;
import org.zkoss.zss.ui.event.EditboxEditingEvent;
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
import org.zkoss.zul.ListModel;
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
	
	@Wire
	Listbox infolist;
	
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		infolist.setModel(infoModel);
		
		addEventInfo("Spreadsheet initialized");
	}

	private void addEventInfo(String info){
		infoModel.add(0, info);
		while(infoModel.size()>100){
			infoModel.remove(infoModel.size()-1);
		}
	}
	
	@Listen("onCellFocused = #ss")
	public void onCellFocused(CellEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Focus on[").append(Ranges.toCellReference(event.getRow(),event.getColumn())).append("]");
		addEventInfo(info.toString());
	}
	
	@Listen("onCellSelection = #ss")
	public void onCellSelection(CellSelectionEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Select on[").append(Ranges.toAreaReference(event.getArea())).append("]");
		addEventInfo(info.toString());
	}
	
	@Listen("onCellSelectionUpdate = #ss")
	public void onCellSelectionUpdate(CellSelectionUpdateEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Selection update from[")
				.append(Ranges.toAreaReference(event.getOrigRow(),event.getOrigColumn(), event.getOrigLastRow(),event.getOrigLastColumn())).append("] to [")
				.append(Ranges.toAreaReference(event.getArea())).append("]");
		addEventInfo(info.toString());
	}

	@Listen("onHeaderUpdate = #ss")
	public void onHeaderUpdate(HeaderUpdateEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Header ").append(event.getAction())
			.append(" on ").append(event.getType());
		switch(event.getType()){
		case COLUMN:
			info.append(" ").append(Ranges.toColumnReference(event.getIndex()));
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
		addEventInfo(info.toString());
	}
	
	
	@Listen("onHeaderClick = #ss")
	public void onHeaderClick(HeaderMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Click on ").append(event.getType()).append(" ");
		
		switch(event.getType()){
		case COLUMN:
			info.append(Ranges.toColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(Ranges.toRowReference(event.getIndex()));
			break;
		}
		addEventInfo(info.toString());
	}
	@Listen("onHeaderRightClick = #ss")
	public void onHeaderRightClick(HeaderMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Right-click on ").append(event.getType()).append(" ");
		
		switch(event.getType()){
		case COLUMN:
			info.append(Ranges.toColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(Ranges.toRowReference(event.getIndex()));
			break;
		}
		addEventInfo(info.toString());
	}
	@Listen("onHeaderDoubleClick = #ss")
	public void onHeaderDoubleClick(HeaderMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Double-click on ").append(event.getType()).append(" ");
		
		switch(event.getType()){
		case COLUMN:
			info.append(Ranges.toColumnReference(event.getIndex()));
			break;
		case ROW:
			info.append(Ranges.toRowReference(event.getIndex()));
			break;
		}
		addEventInfo(info.toString());
	}
	
	@Listen("onCellClick = #ss")
	public void onCellClick(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Click on cell ").append(Ranges.toCellReference(event.getRow(),event.getColumn()));
		
		addEventInfo(info.toString());
	}
	@Listen("onCellRightClick = #ss")
	public void onCellRightClick(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Right-click on cell ").append(Ranges.toCellReference(event.getRow(),event.getColumn()));
		
		addEventInfo(info.toString());
	}
	@Listen("onCellDoubleClick = #ss")
	public void onCellDoubleClick(CellMouseEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Double-click on cell ").append(Ranges.toCellReference(event.getRow(),event.getColumn()));
		
		addEventInfo(info.toString());
	}
	
	@Listen("onStartEditing = #ss")
	public void onStartEditing(StartEditingEvent event){
		StringBuilder info = new StringBuilder();
		String ref = Ranges.toCellReference(event.getRow(),event.getColumn());
		info.append("Start to edit ").append(ref)
		.append(", editing-value is ").append(event.getEditingValue()).append(" client-value is ").append(event.getClientValue());
		addEventInfo(info.toString());
		
		if(ref.equals("D1")){
			String newval = "Supprise!!";
			//we change the editing value
			event.setEditingValue(newval);
			addEventInfo("Editing value is change to "+newval);
		}else if(ref.equals("E1")){
			//we don't allow edit D2
			event.cancel();
			addEventInfo("Editing is canceled");
		}else{
			addEventInfo("Try to edit D1 or E1");
		}
	}
	
	@Listen("onStopEditing = #ss")
	public void onStopEditing(StopEditingEvent event){
		StringBuilder info = new StringBuilder();
		String ref = Ranges.toCellReference(event.getRow(),event.getColumn());
		info.append("Stop editing ").append(ref)
		.append(", value is ").append(event.getEditingValue());
		addEventInfo(info.toString());
		
		if(ref.equals("D3")){
			String newval = event.getEditingValue()+"-Woo";
			//we change the editing value
			event.setEditingValue(newval);
			addEventInfo("Editing value is change to "+newval);
		}else if(ref.equals("E3")){
			//we don't allow edit D2
			event.cancel();
			addEventInfo("Editing is canceled");
		}else{
			addEventInfo("Try to edit D3 or E3");
		}
	}
	
	@Listen("onEditboxEditing = #ss")
	public void onEditboxEditing(EditboxEditingEvent event){
		StringBuilder info = new StringBuilder();
		String ref = Ranges.toCellReference(event.getRow(),event.getColumn());
		info.append("Editing ").append(ref)
		.append(", value is ").append(event.getEditingValue());
		addEventInfo(info.toString());
	}
	
	@Listen("onSheetSelect = #ss")
	public void onSheetSelect(SheetSelectEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Select sheet : ").append(event.getSheetName());
		addEventInfo(info.toString());
	}
	
	@Listen("onSheetNameChange = #ss")
	public void onSheetNameChange(SheetEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Rename sheet to ").append(event.getSheetName());
		addEventInfo(info.toString());
	}
	
	@Listen("onSheetOrderChange = #ss")
	public void onSheetOrderChange(SheetEvent event){
		StringBuilder info = new StringBuilder();
		Sheet sheet = event.getSheet();
		info.append("Reorder sheet : ").append(event.getSheetName()).append(" to ").append(sheet.getBook().getSheetIndex(sheet));
		addEventInfo(info.toString());
	}
	
	@Listen("onSheetCreate = #ss")
	public void onSheetCreate(SheetEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Create sheet : ").append(event.getSheetName());
		addEventInfo(info.toString());
	}
	
	@Listen("onSheetDelete = #ss")
	public void onSheetDelete(SheetDeleteEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Delete sheet : ").append(event.getSheetName());
		addEventInfo(info.toString());
	}
	
	@Listen("onCtrlKey = #ss")
	public void onCtrlKey(KeyEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Keys : ").append(event.getKeyCode()).append(", ctrl:").append(event.isCtrlKey())
		.append(", alt:").append(event.isAltKey()).append(", shift:").append(event.isShiftKey());
		addEventInfo(info.toString());
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
				.append(Ranges.toAreaReference(anchor.getRow(),
						anchor.getColumn(), anchor.getLastRow(),
						anchor.getLastColumn()));
		addEventInfo(info.toString());
	}
	
	@Listen("onWidgetCtrlKey = #ss")
	public void onWidgetCtrlKey(WidgetKeyEvent event){
		StringBuilder info = new StringBuilder();
		info.append("Widget ").append(event.getWidgetData()).append(" Key : ").append(event.getKeyCode()).append(", ctrl:").append(event.isCtrlKey())
		.append(", alt:").append(event.isAltKey()).append(", shift:").append(event.isShiftKey());
		addEventInfo(info.toString());
	}
	
	
	/**
	 * ON_CTRL_KEY
	 * ON_AUX_ACTION
	 * ON_HYPERLINK
	 * ON_CELL_FILTER //useless
	 * ON_CELL_VALIDATOR //useless
	 * 
	 * ==Widget Events==
	 * ON_WIDGET_CTRL_KEY
	 * ON_WIDGET_UPDATE
	 * 
	 * ==BookEvents==
	 * ON_SHEET_NAME_CHANGE
	 * ON_SHEET_ORDER_CHANGE
	 * ON_SHEET_DELETE
	 * ON_SHEET_CREATE
	 * ON_CELL_CHANGE
	 * 
	 * 
	 */
}



