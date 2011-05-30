/* FormulaEditor.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 4:28:32 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import java.util.HashSet;
import java.util.LinkedHashSet;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.InputEvent;
import org.zkoss.zk.ui.event.KeyEvent;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.WorkbookCtrl;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.SheetCtrl;
import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.EditboxEditingEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.StartEditingEvent;
import org.zkoss.zss.ui.event.StopEditingEvent;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.impl.XulElement;

/**
 * @author Sam
 *
 */
public class FormulaEditor extends Textbox {
	private final static Character[] KEY = 
		new Character[]{'=', '+', '-', '*', '/', '!', ':', '^', '&', '(',  ',', '.'};
	private final static HashSet<Character> Separator =
		new HashSet<Character>();
	static {
		addClientEvent(FormulaEditor.class, "onTab", CE_REPEAT_IGNORE);
		for (int i = 0; i < KEY.length; i++) {
			Separator.add(KEY[i]);
		}
	}

	private String oldEdit;
	private String oldText;
	private String newEdit;
	
	private boolean everFocusCell = false;
	private boolean focusOut = false;
	private Cell currentEditcell;
	private Boolean everTab = null;

	private int formulaRow;
	private int formulaColumn;
	private Cell formulaCell;
	/*cache added focus names*/
	private LinkedHashSet<String> addedFocusNames = new LinkedHashSet<String>();
	private Worksheet formulaSheet;
	private String focusCellRef;
	/**
	 * Indicate edit existing formula.
	 */
	private boolean editExistingFormula = false;

	private final static String FORMULA_FOCUS_NAME = "currentFormulaFocus";
	private final static String FORMULA_COLOR = "#555555";
	private final static String[] FORMULA_COLORS = 
		new String[]{"#0000FF", "#008000", "#9900CC", "#800000", "#800000", "#FF6600", "#CC0099"};
	
	public FormulaEditor() {
		setMultiline(true);
	}

	private void addFormulaEditCellFocus() {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		bookCtrl.moveEditorFocus(FORMULA_FOCUS_NAME, FORMULA_COLOR, formulaRow, formulaColumn);
		addedFocusNames.add(FORMULA_FOCUS_NAME);
	}
	
	/**
	 * Generate cell focus from edit label
	 */
	private void generateCellFocus(String edit) {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		boolean startEditingFormula = addedFocusNames.size() == 0;
		int maxCol = bookCtrl.getMaxcolumns();
		int maxRow = bookCtrl.getMaxrows();
		
		int currentColumn = bookCtrl.getSelection().getLeft();
		int currentRow = bookCtrl.getSelection().getTop();
		
		LinkedHashSet<String> currentFocus = new LinkedHashSet<String>(addedFocusNames);
		LinkedHashSet<String> newFocus = new LinkedHashSet<String>();
		if (addedFocusNames.contains(FORMULA_FOCUS_NAME))
			newFocus.add(FORMULA_FOCUS_NAME);

		LinkedHashSet<String> cellRefs = parseCellReference(edit);
		for (String name : cellRefs) {
			if (addedFocusNames.contains(name) || "".equals(name)) {
				newFocus.add(name);
				continue;
			}
			Range rng = null;
			try {
				rng = Ranges.range(formulaSheet, name);
			} catch (NullPointerException ex) { /*input wrong cell reference will cause NPE or IllegalArgumentException*/
			} catch (IllegalArgumentException ex) {
			}
			if (rng == null)
				continue;
			int col = rng.getColumn();
			int row = rng.getRow();
			if (col >= maxCol || row >= maxRow || 
				(col == formulaColumn && row == formulaRow) ||
				(col == currentColumn && row == currentRow))
				continue;

			String color = FORMULA_COLORS[
			               (addedFocusNames.contains(FORMULA_FOCUS_NAME) ? addedFocusNames.size() - 1 : addedFocusNames.size()) % FORMULA_COLORS.length];
			newFocus.add(name);
			addedFocusNames.add(name);
			bookCtrl.moveEditorFocus(name, color, row, col);
		}
		if (startEditingFormula) {
			newFocus.add(FORMULA_FOCUS_NAME);
			addFormulaEditCellFocus();
		}
		removeNonexistentFocus(currentFocus, newFocus);
	}
	
	private void removeNonexistentFocus(LinkedHashSet<String> currentFocus, LinkedHashSet<String> newFocus) {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		for (String ref : currentFocus) {
			if (!newFocus.contains(ref)) {
				bookCtrl.removeEditorFocus(ref);
				addedFocusNames.remove(ref);
			}
		}
	}
	
	private LinkedHashSet<String> parseCellReference(String edit) {
		int startIdx = -1;
		int endIdx = -1;
		//boolean find = false;
		LinkedHashSet<String> val = new LinkedHashSet<String>();
		for (int i = 0; i < edit.length(); i++) {
			if (Separator.contains(edit.charAt(i))) {
				startIdx = i;
				for (int j = i + 1; j < edit.length(); j++) {
					if (Separator.contains(edit.charAt(j))) {
						//find = true;
						endIdx = j;
						val.add(edit.substring(i + 1, j));
						break;
					}
				}
				if (endIdx > startIdx) {
					i = startIdx = endIdx - 1;
					endIdx = -1;
					//find = false;
				} else {
					val.add(edit.substring(i + 1));
				}
			}
		}
		return val;
	}
	
	private void cacheFormulaEditingInfo() {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		formulaRow = bookCtrl.getSelection().getTop();
		formulaColumn = bookCtrl.getSelection().getLeft();
		formulaCell = Utils.getCell(bookCtrl.getSelectedSheet(), formulaRow, formulaColumn);
		formulaSheet = bookCtrl.getSelectedSheet();
	}
	
	private boolean isStartEditingFormula(String edit) {
		if ("=".equals(edit))
			return true;
		if (edit != null && edit.startsWith("=") && addedFocusNames.size() == 0)
			return true;
		return false;
	}
	
	public void onChanging(InputEvent event) {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		newEdit = event.getValue();
		if (isComposingFormula(newEdit)) {
			//if (editExistingFormula && allowAppendCellReference(newEdit))
			//	editExistingFormula = allowAppendCellReference(newEdit);
			bookCtrl.escapeAndUpdateText(formulaCell, newEdit);
			if (isStartEditingFormula(newEdit))
				cacheFormulaEditingInfo();
			generateCellFocus(newEdit);
		}  else if (currentEditcell != null) {
			final int left = bookCtrl.getSelection().getLeft();
			final int top = bookCtrl.getSelection().getTop();
			final Worksheet sheet = bookCtrl.getSelectedSheet();
			currentEditcell = Utils.getOrCreateCell(sheet, top, left);
			bookCtrl.escapeAndUpdateText(currentEditcell, newEdit);
		}
	}

	private static boolean allowAppendCellReference(String editLabel) {
		return editLabel != null && editLabel.length() > 0 &&
				Separator.contains(editLabel.charAt(editLabel.length() - 1));
	}

	public void onCancel() {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		focusOut = true;
		recoverEditorText();
		recoverCellText();
		endEditingFormula(false);
		int row = bookCtrl.getSelection().getTop();
		int col = bookCtrl.getSelection().getLeft();
		bookCtrl.focusTo(row, col, true);
	}
	
	public void onFocus() {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		Worksheet sheet = bookCtrl.getSelectedSheet();
		if (sheet == null) { //no sheet, no operation
			return;
		}
		bookCtrl.reGainFocus();

		//TODO: why need it ?
		//newEdit = null;

		everFocusCell = false;
		everTab = null;
		focusOut = false;
		int left = bookCtrl.getSelection().getLeft();
		int top = bookCtrl.getSelection().getTop();
		currentEditcell = Utils.getCell(sheet, top, left);
		
		if (currentEditcell != null) {
			oldEdit = Ranges.range(sheet, top, left).getEditText();
			oldText = Utils.getCellText(sheet, currentEditcell); //escaped HTML to show cell value
			if (!isComposingFormula(newEdit))
				bookCtrl.escapeAndUpdateText(currentEditcell, oldEdit);

			if (!editExistingFormula && isComposingFormula(oldEdit) && addedFocusNames.size() == 0) {
				editExistingFormula = true;
				cacheFormulaEditingInfo();
				generateCellFocus(oldEdit);
			}
		}
		bookCtrl.clearClipbook();
	}
	
	public void onTab(Event event) {
		everTab = Boolean.valueOf(((KeyEvent)event).isShiftKey());
	}
	
	public void onChange(Event event) {
		newEdit = ((InputEvent)event).getValue(); //remember the changed value
	}
	
	private void handleCellText() {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		if (newEdit == null) { //no change
			recoverCellText(); //recover cell text
		} else if (currentEditcell != null){
			if (!isComposingFormula(newEdit))
				Utils.setEditText(bookCtrl.getSelectedSheet(), 
						currentEditcell.getRowIndex(), currentEditcell.getColumnIndex(), newEdit);
		}
	}
	
	public void onBlur() {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		if (focusOut) { //onChange, onOK, or onCancel already done everything!
			return;
		}
		if (editExistingFormula) {
			if (!allowAppendCellReference(newEdit)) {
				endEditingFormula(false);
			}
			editExistingFormula = false;
		}
		focusOut = true;
		
		handleCellText();
		
		Position pos = bookCtrl.getCellFocus();
		int row = pos.getRow();
		int col = pos.getColumn();
		int oldrow = currentEditcell == null ? -1 : currentEditcell.getRowIndex();
		int oldcol = currentEditcell == null ? -1 : currentEditcell.getColumnIndex();
		final boolean focusChanged = (row != oldrow || col != oldcol); //user click directly to a different cell
		if (!focusChanged) {
			if (everTab != null) { //Tab key
				final Worksheet sheet = bookCtrl.getSelectedSheet();
				if (everTab.booleanValue()) { //shift + Tab
					col = col - 1;
					if (col < 0) {
						col = 0;
					} else {
						final CellRangeAddress merged = sheet != null ? ((SheetCtrl)sheet).getMerged(row, col) : null;
						if (merged != null) {
							col = merged.getFirstColumn();
						}
					}
				} else { //Tab
					final CellRangeAddress merged = sheet != null ? ((SheetCtrl)sheet).getMerged(row, col) : null;
					col = merged == null ? col + 1 : merged.getLastColumn() + 1;
					if (bookCtrl.getMaxcolumns() <= col) {
						col = bookCtrl.getMaxcolumns() - 1;
						final CellRangeAddress newmerged = sheet != null ? ((SheetCtrl)sheet).getMerged(row, col) : null;
						if (newmerged != null) {
							col = newmerged.getFirstColumn();
						}
					}
				}
				bookCtrl.focusTo(row, col, true);
			} else if (everFocusCell) { //click on the same cell, shall enter edit mode, something like press F2
				//TODO click on the same cell, shall enter edit mode, something like press F2
			}
		}
	}
	
	private void recoverEditorText() {
		setText(oldEdit);
	}
	private void recoverCellText() {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		if (oldText != null && currentEditcell != null) {
			//already escape, simply update
			bookCtrl.updateText(currentEditcell, oldText);
			bookCtrl.escapeAndUpdateText(formulaCell, oldText);
		}
	}

	public void onOK() {
		if (formulaCell != null) {
			endEditingFormula(true);
		}
		WorkbookCtrl bookCtrl  = getDesktopWorkbenchContext().getWorkbookCtrl();		
		focusOut = true;
		handleCellText();
		
		//move cell focus
		Position pos = bookCtrl.getCellFocus();
		int row = pos.getRow() + 1;
		int col = pos.getColumn();
		if (bookCtrl.getMaxrows() <= row) {
			row = bookCtrl.getMaxrows() - 1;
		}
		bookCtrl.focusTo(row, col, true);
	}

	private static boolean isComposingFormula(String label) {
		return label != null && label.startsWith("=");
	}

	/**
	 * 
	 * @param target
	 * @param replacement
	 */
	private boolean replaceFormulaReference(String target, String replacement) {
		final WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		int refStartIndex = -1;
		for (int i = newEdit.length() - 1; i >= 0; i--) {
			char c = newEdit.charAt(i);
			if (Separator.contains(c)) {
				refStartIndex = i;
				break;
			}
		}

		if (refStartIndex >= 0 && target != null && !target.equals(replacement) &&
				(newEdit.indexOf(target, refStartIndex) + target.length() ==  newEdit.length())) {
			newEdit = newEdit.substring(0, refStartIndex + 1) + replacement;
			setText(newEdit);
			bookCtrl.escapeAndUpdateText(formulaCell, newEdit);
			return true;
		}
		return false;
	}

	private void appendFormulaReference(String cellRef) {
		final WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		if (allowAppendCellReference(newEdit)) {
			newEdit += focusCellRef;
			setText(newEdit);
			bookCtrl.escapeAndUpdateText(formulaCell, newEdit);
		}
	}
	
	private void clearCellReferenceFocus() {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		for (String name : addedFocusNames) {
			bookCtrl.removeEditorFocus(name);
		}
		addedFocusNames.clear();
	}
	
	private void endEditingFormula(boolean confirmChange) {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		
		Range range = Ranges.range(formulaSheet, formulaCell.getRowIndex(), formulaCell.getColumnIndex());
		if (confirmChange) {
			range.setEditText(newEdit);
			bookCtrl.focusTo(formulaRow, formulaColumn, false);
		}
		clearCellReferenceFocus();
		editExistingFormula = false;
		formulaCell = null;
		newEdit = null;
		formulaRow = 0;
		formulaColumn = 0;
	}

	public void onCreate() {
		final WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		bookCtrl.addEventListener(Events.ON_CELL_FOUCSED, new EventListener() {
			@Override
			public void onEvent(Event event) throws Exception {
				CellEvent cellEvent = (CellEvent)event;
				everFocusCell = true;

				if (!isComposingFormula(newEdit)) {
					Cell cell = Utils.getCell(cellEvent.getSheet(), cellEvent.getRow(), cellEvent.getColumn());
					String editText = Utils.getEditText(cell);
					setText(cell == null ? "" : (editText == null ? "" : editText));
				} else {
					if (!editExistingFormula) {
						FormulaEditor.this.focus();
						String lastFocusRef = focusCellRef;
						focusCellRef = bookCtrl.getReference(cellEvent.getRow(), cellEvent.getColumn());
						if (replaceFormulaReference(lastFocusRef, focusCellRef)) 
							return;
						else
							appendFormulaReference(focusCellRef);
						generateCellFocus(newEdit);
					}
				}	
			}
		});

		bookCtrl.addEventListener(Events.ON_STOP_EDITING,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						StopEditingEvent evt = (StopEditingEvent)event;
						//setText((String) evt.getEditingValue());

						//chart not implement yet
//						// to notify all widgets there is a cell changed
//						for (int i = 0; i < chartKey; i++) {
//							try {
//								Window win = (Window) mainWin.getFellow("chartWin" + i);
//								if (win != null) {
//									Chart myChart = (Chart) win.getFellow("myChart");
//									CellEvent event = new StopEditingEvent(
//											org.zkoss.zss.ui.event.Events.ON_STOP_EDITING,
//											myChart, evt.getSheet(), evt.getRow(), evt
//													.getColumn(), (String) evt.getData());
//									org.zkoss.zk.ui.event.Events.postEvent(event);
//								}
//							} catch (Exception e) {
//								// the chart may be deleted
//							}
//						}
					}
				});
		//TODO: add ON_START_EDITING, add cell focus if cell is formula
		bookCtrl.addEventListener(Events.ON_START_EDITING,
				new EventListener() {
			public void onEvent(Event event) throws Exception {
				StartEditingEvent evt = (StartEditingEvent)event;
				newEdit = (String) evt.getClientValue();
				if (isStartEditingFormula(newEdit)) {
					cacheFormulaEditingInfo();
				}
				if (isComposingFormula(newEdit)) {
					int left = bookCtrl.getSelection().getLeft();
					int top = bookCtrl.getSelection().getTop();
					formulaCell = Utils.getCell(bookCtrl.getSelectedSheet(), top, left);

					addedFocusNames.add(FORMULA_FOCUS_NAME);
					bookCtrl.moveEditorFocus(FORMULA_FOCUS_NAME, FORMULA_COLOR, top, left);
				}
				setText((String) evt.getEditingValue());
			}
		});
		bookCtrl.addEventListener(Events.ON_EDITBOX_EDITING,
				new EventListener() {
			public void onEvent(Event event) throws Exception {
				EditboxEditingEvent evt = (EditboxEditingEvent)event;
				setText((String) evt.getEditingValue());
			}
		});
		final DesktopWorkbenchContext cnt = getDesktopWorkbenchContext();
		cnt.addEventListener(Consts.ON_SHEET_CHANGED, new EventListener() {
			public void onEvent(Event event) throws Exception {
				//TODO: provide insert sheet reference to formula
				if (isComposingFormula(newEdit)) {
					endEditingFormula(true);
				}
			}
		});
	}

	private DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(this);
	}

	/** Processes an AU request.
	 *
	 * <p>Default: in addition to what are handled by {@link XulElement#service},
	 * it also handles onChange, onChanging and onError.
	 * @since 2.0.0
	 */
	public void service(org.zkoss.zk.au.AuRequest request, boolean everError) {
		final String cmd = request.getCommand();
		if (cmd.equals("onTab")) {
			org.zkoss.zk.ui.event.Events.postEvent(KeyEvent.getKeyEvent(request));
		} else
			super.service(request, everError);
	}
}