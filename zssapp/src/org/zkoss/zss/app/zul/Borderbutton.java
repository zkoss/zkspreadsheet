package org.zkoss.zss.app.zul;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zss.app.Dropdownbutton;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;

public class Borderbutton extends Dropdownbutton implements ZssappComponent, IdSpace {

	private final static String URI = "~./zssapp/html/button/borderSelector.zul";
	
	private Spreadsheet ss;
	
	private boolean skipClickEvent;
	
	public Borderbutton() {
		
		Executions.createComponents(URI, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
		
		setImage("/img/border-bottom.png");
		setTooltip(Labels.getLabel("border.bottom"));
	}
	
	public void onClick(Event event) {
		if (skipClickEvent) {
			skipClickEvent = !skipClickEvent;
			return;
		}
		
		Utils.setBorder(ss.getSelectedSheet(),
				ss.getSelection(), 
				getBorderType(Labels.getLabel("border.bottom")), 
				BorderStyle.MEDIUM, 
				"#000000");
	}
	
	public void onBorderSelector(ForwardEvent evt) {
		final String color = "#000000";

		String param = (String)evt.getData();
		
		if (param.equals(Labels.getLabel("border.no")))
			Utils.setBorder(ss.getSelectedSheet(), ss.getSelection(), BookHelper.BORDER_FULL, BorderStyle.NONE, color);
		else
			Utils.setBorder(ss.getSelectedSheet(), ss.getSelection(), getBorderType(param), BorderStyle.MEDIUM, color);
		
		skipClickEvent = true;
		Events.postEvent(Events.ON_CLICK, this, null);
	}
	
	/**
	 * Returns the border type base on i3-label
	 * <p>
	 * Default: returns {@link #BookHelper.BORDER_EDGE_BOTTOM}, if no match.
	 * 
	 * @param i3label
	 */
	public static short getBorderType(String i3label) {
		if (i3label == null || i3label.equals(Labels.getLabel("border.bottom"))) {
			return BookHelper.BORDER_EDGE_BOTTOM;
		}

		if (i3label.equals(Labels.getLabel("border.bottom")))
			return BookHelper.BORDER_EDGE_BOTTOM;
		else if (i3label.equals(Labels.getLabel("border.right")))
			return BookHelper.BORDER_EDGE_RIGHT;
		else if (i3label.equals(Labels.getLabel("border.top")))
			return BookHelper.BORDER_EDGE_TOP;
		else if (i3label.equals(Labels.getLabel("border.left")))
			return BookHelper.BORDER_EDGE_LEFT;
		else if (i3label.equals(Labels.getLabel("border.insideHorizontal")))
			return BookHelper.BORDER_INSIDE_HORIZONTAL;
		else if (i3label.equals(Labels.getLabel("border.insideVertical")))
			return BookHelper.BORDER_INSIDE_VERTICAL;
		else if (i3label.equals(Labels.getLabel("border.diagonalDown")))
			return BookHelper.BORDER_DIAGONAL_DOWN;
		else if (i3label.equals(Labels.getLabel("border.diagonalUp")))
			return BookHelper.BORDER_DIAGONAL_UP;
		else if (i3label.equals(Labels.getLabel("border.full")))
			return BookHelper.BORDER_FULL;
		else if (i3label.equals(Labels.getLabel("border.outside")))
			return BookHelper.BORDER_OUTLINE;
		else if (i3label.equals(Labels.getLabel("border.inside")))
			return BookHelper.BORDER_INSIDE;
		else if (i3label.equals(Labels.getLabel("border.diagonal")))
			return BookHelper.BORDER_DIAGONAL;

		return BookHelper.BORDER_EDGE_BOTTOM;
	}

	@Override
	public Spreadsheet getSpreadsheet() {
		return ss;
	}

	@Override
	public void setSpreadsheet(Spreadsheet spreadsheet) {
		ss = checkNotNull(spreadsheet, "Spreadsheet is null");
		
		setWidgetListener("onClick", "this.$f('" + ss.getId() + "', true).focus(false);");
		setWidgetListener("onDropdown", "this.$f('" + ss.getId() + "', true).focus(false);");
	}	
}
