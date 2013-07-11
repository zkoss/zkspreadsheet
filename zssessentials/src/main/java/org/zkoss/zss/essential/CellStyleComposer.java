package org.zkoss.zss.essential;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.CellStyle.VerticalAlignment;
import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Label;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.ext.Selectable;

/**
 * Demonstrate CellStyle, CellOperationUtil API usage
 * 
 * @author Hawk
 * 
 */
public class CellStyleComposer extends SelectorComposer<Component> {

	@Wire
	private Label cellRef;
	@Wire
	private Label hAlign;
	@Wire
	private Label vAlign;
	@Wire
	private Label tBorder;
	@Wire
	private Label bBorder;
	@Wire
	private Label lBorder;
	@Wire
	private Label rBorder;
	@Wire
	private Listbox hAlignBox;
	@Wire
	private Listbox vAlignBox;
	@Wire
	private Spreadsheet ss;

	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		hAlignBox.setModel(getAlignmentList());
		vAlignBox.setModel(getVertocalAlignmentList());
	}

	@Listen("onCellFocused = #ss")
	public void onCellFocused() {
		Position pos = ss.getCellFocus();
		refreshCellStyle(pos.getRow(), pos.getColumn());
	}

	private void refreshCellStyle(int row, int col) {
		Range range = Ranges.range(ss.getSelectedSheet(), row, col);

		cellRef.setValue(Ranges.getCellReference(row, col));

		CellStyle style = range.getCellStyle();

		// display cell style
		hAlign.setValue(style.getAlignment().name());
		vAlign.setValue(style.getVerticalAlignment().name());
		tBorder.setValue(style.getBorderTop().name());
		bBorder.setValue(style.getBorderBottom().name());
		lBorder.setValue(style.getBorderLeft().name());
		rBorder.setValue(style.getBorderRight().name());

		// update to editor
		List<Object> selection = new ArrayList<Object>();
		selection.add(style.getAlignment());
		((Selectable) hAlignBox.getModel()).setSelection(selection);
		selection.clear();
		selection.add(style.getVerticalAlignment());
		((Selectable) vAlignBox.getModel()).setSelection(selection);

	}

	@Listen("onSelect = #hAlignBox")
	public void applyAlignment() {

		Range selection = Ranges.range(ss.getSelectedSheet(), ss.getSelection());
		CellOperationUtil.applyAlignment(selection
				, (Alignment)hAlignBox.getSelectedItem().getValue());
	}

	@Listen("onSelect = #vAlignBox")
	public void applyVerticalAlignment() {

		Range selection = Ranges.range(ss.getSelectedSheet(), ss.getSelection());
		CellOperationUtil.applyVerticalAlignment(selection
				, (VerticalAlignment)vAlignBox.getSelectedItem().getValue());
	}

	private ListModelList<Alignment> getAlignmentList() {
		ListModelList<Alignment> list = new ListModelList<Alignment>();
		list.add(Alignment.LEFT);
		list.add(Alignment.CENTER);
		list.add(Alignment.RIGHT);
		list.add(Alignment.GENERAL);

		return list;
	}

	private ListModelList<VerticalAlignment> getVertocalAlignmentList() {
		return new ListModelList<VerticalAlignment>(Arrays.asList(VerticalAlignment
				.values()));
	}

}
