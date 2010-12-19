/* CustomSortWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Aug 17, 2010 6:34:44 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.sort;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.zkoss.lang.Objects;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.SelectEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Button;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Comboitem;
import org.zkoss.zul.Label;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listcell;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.ListitemRenderer;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class CustomSortWindowCtrl extends GenericForwardComposer {
	/**
	 * Sort by column
	 */
	private final static boolean SORT_TOP_TO_BOTTOM = false;
	/**
	 * Sort by row
	 */
	private final static boolean SORT_LEFT_TO_RIGHT = true;
	

	private final static String ORIENTATION_COLUMN_KEY = "sort.orientation.column";// orientation

	/**
	 * Indicate the sort orientation
	 * <p> Default: false, sort from top to bottom, means sort by column
	 * <p> true, sort from left to right, means sort by row
	 */
	private boolean sortOrientation;
	
	/**
	 * Selected sort level target by user
	 */
	private ListModelList sortLevelModel;
	
	/**
	 * Depends on sortOrientation
	 * If sort by rows, available sort target include top row to bottom row
	 * If sort by columns, available sort target include left column to right column
	 * 
	 * The value build from spreadsheet's selection range.
	 */
	private List<String> availableSortIndex = new ArrayList<String>();
	private ListModelList sortIndexModel = new ListModelList();
	
	private Listbox sortLevel;
	private Spreadsheet ss;
	private Checkbox caseSensitive;
	private Checkbox hasHeader;
	private Listbox sortOrientationLB;
	private Window sortWin;
	private Button addBtn;
	private Button delBtn;
	private Button upBtn;
	private Button downBtn;
	private Button okBtn;
	
	
	public CustomSortWindowCtrl () {
		ss = (Spreadsheet)getParam("spreadsheet");
		if (ss == null)
			throw new UiException("Spreadsheet object is empty");
	} 
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		initSortLevelListbox();
	}

	private void initSortLevelListbox () {
		List<SortLevel> ary = new ArrayList<SortLevel>();
		ary.add(new SortLevel());
		setAvailableSortTarget(availableSortIndex);
		
		sortLevelModel = new ListModelList(ary);
		sortLevel.setModel(sortLevelModel);
		sortLevel.setItemRenderer(sortLevelRenderer);
	}
	
	/**
	 * Setup the available sort target
	 * <p> Sort target is column if sort direction is from top to bottom
	 * <p> Sort target is row if sort direction is from left to right
	 */
	private void setAvailableSortTarget (List<String> list) {
		list.clear();
		list.add(new String(""));
		Rect rect = ss.getSelection();
		if (sortOrientation == SORT_TOP_TO_BOTTOM) {
			for (int i = rect.getLeft(); i <= rect.getRight(); i++)
				list.add(new String("Column " + ss.getColumntitle(i)));
		} else {
			for (int i = rect.getTop(); i <= rect.getBottom(); i++)
				list.add(new String("Row " + ss.getRowtitle(i)));
		}
		sortIndexModel.clear();
		sortIndexModel.addAll(list);
	}
	
	public void onSelect$sortLevel () {
		setClickableMoveButtons();
	}
	
	public void onInitRender$sortLevel () {
		setClickableMoveButtons();
	}
	
	private void setClickableMoveButtons() {
		int idx = sortLevel.getSelectedIndex();
		upBtn.setDisabled(idx == 0 ? true : false);
		downBtn.setDisabled(idx < ( sortLevel.getItemCount() - 1) ? false : true);
	}
	
	public void onClick$addBtn () {
		sortLevelModel.add(new SortLevel());
	}
	
	public void onClick$delBtn () {
		sortLevelModel.remove(sortLevel.getSelectedIndex());
	}
	
	public void onClick$upBtn () {
		int idx = sortLevel.getSelectedIndex();
		int forward = idx - 1;
		if (forward >= 0) {
			swapSortLevel(sortLevelModel, idx, forward);
			sortLevel.setSelectedIndex(forward);
		}
	}
	
	public void onClick$downBtn () {
		int idx = sortLevel.getSelectedIndex();
		int back = idx + 1;
		if (back < sortLevelModel.getSize()) {
			swapSortLevel(sortLevelModel, idx, back);
			sortLevel.setSelectedIndex(back);
		}
	}
	
	public void onClick$okBtn () {
		if (hasEmptyArgs(sortLevelModel.getInnerList())) {
			try {
				Messagebox.show(getLabel("sort.err.hasEmptyField"));
			} catch (InterruptedException e) {
			}
			return;
		}
		
		String dupTarget = checkDuplicateSortIndex(sortLevelModel.getInnerList());
		if (dupTarget != null) {
			try {
				Messagebox.show(dupTarget + " " + getLabel("sort.err.duplicateField"));
			} catch (InterruptedException e) {
			}
			return;
		}
		
		int count = 0;
		int size = sortLevelModel.size() <= 3 ? sortLevelModel.size() : 3;
		int[] index = new int[size];
		int[] dataOption = new int[size];
		boolean[] algorithm = new boolean[size];
		for (Iterator iter = sortLevelModel.iterator(); iter.hasNext() && count < index.length; count++) {
			SortLevel l = (SortLevel)iter.next();
			index[count] = l.sortIndex;
			algorithm[count] = l.order;
			dataOption[count] = l.dataOption;
		}
		
		//call utils
		Utils.sort(ss.getSelectedSheet(), ss.getSelection(), index, algorithm, dataOption,  hasHeader.isChecked(), caseSensitive.isChecked(), sortOrientation);
		sortWin.detach();
	}
	
	/**
	 * Returns true if there is a
	 * @return
	 */
	private static boolean hasEmptyArgs (List<SortLevel> list) {
		for (SortLevel s : list) {
			if (s.sortIndex == -1)
				return true;
		}
		return false;
	}
	
	/**
	 * Returns the duplicate sort title
	 * @param list
	 * @return string if there is a duplicate, or null when there is no duplicate.
	 */
	private String checkDuplicateSortIndex(List<SortLevel> list) {
		HashMap<Integer, Boolean> map = new HashMap<Integer, Boolean>();
		for (SortLevel s : list) {
			if ( map.containsKey( Integer.valueOf(s.sortIndex) ) ) {
				String title = null;
				if (sortOrientation) 
					title = "Row " + ss.getRowtitle(s.sortIndex);
				else
					title = "Column " + ss.getColumntitle(s.sortIndex);
				return title;
			} else
				map.put(Integer.valueOf(s.sortIndex), Boolean.valueOf(true));
		}
		return null;
	}
	

	
	private static void swapSortLevel(ListModelList model, int sourceIdx, int destIdx) {
		SortLevel srcSort = (SortLevel)model.get(sourceIdx);
		SortLevel dstSort = (SortLevel)model.get(destIdx);
		model.set(destIdx, srcSort);
		model.set(sourceIdx, dstSort);
	}
	
	public void onSelect$sortOrientationLB () {
		boolean orientation;
		Listitem seld = sortOrientationLB.getSelectedItem();
		orientation = getLabel(ORIENTATION_COLUMN_KEY).equals(seld) ? 
				SORT_TOP_TO_BOTTOM : SORT_LEFT_TO_RIGHT;
		if (sortOrientation != orientation) {
			sortOrientation = orientation;
			setAvailableSortTarget(availableSortIndex);
		}
	}
	
	private static Object getParam (String key) {
		return Executions.getCurrent().getArg().get(key);
	}
	
	private ListitemRenderer sortLevelRenderer = new ListitemRenderer() {

		public void render(Listitem item, Object obj) throws Exception {
			Listcell cell = new Listcell();
			SortLevel sort = (SortLevel)obj;
			SortLevel first = (SortLevel)sortLevelModel.get(0);
			cell.appendChild(sort.equals(first) ? 
					(Component)new Label(getLabel("sort.sortBy")) : (Component)new Label(getLabel("sort.thenBy")));
			item.appendChild(cell);
			
			cell = new Listcell();
			SortIndexSelector idxSel = new SortIndexSelector(sort);
			cell.appendChild(idxSel);
			item.appendChild(cell);
			
			cell = new Listcell();
			SortAlgorithm sortMethod = new SortAlgorithm(sort);
			cell.appendChild(sortMethod);
			idxSel.setAttribute(SortAlgorithm.class.getCanonicalName(), sortMethod);
			item.appendChild(cell);
		}
		
	};
	
	
	public class SortIndexSelector extends Combobox {
		SortLevel sort;
		
		public SortIndexSelector (SortLevel sort) {
			setWidth("100%");
			setReadonly(true);
			setMold("rounded");
			this.sort = sort;
			setModel(sortIndexModel);
		} 
		
		public void onAfterRender() {
			if (sort.sortIndex >= 0 && getItemCount() > 0) {
				int idx = getSpreadsheetIndexOffset(ss, sort.sortIndex, sortOrientation);
				setSelectedIndex(idx >= 0 ? idx : 0);
			}
		}

		public void onSelect () {
			String seldLabel = getSelectedItem().getLabel();
			if (seldLabel != null && seldLabel.length() > 0) {
				int blankIdx = seldLabel.lastIndexOf(" ");
				if (blankIdx >= 0) {
					String title = seldLabel.substring( seldLabel.lastIndexOf(" ") + " ".length() );
					int idx = getSpreadsheetIndexBy(ss, title, sortOrientation);
					if (idx != -1) {
						sort.sortIndex = idx;
						SortAlgorithm alg = (SortAlgorithm)getAttribute(SortAlgorithm.class.getCanonicalName());
						alg.refresh();
					}
				}
			}
		}
	}
	
	public class SortAlgorithm extends Combobox {
		private SortLevel sort;
		
		/**
		 * Default display the String 
		 */
		private final static String STR_ASCENDING_KEY = "sort.str.ascending";
		private final static String STR_DESCENDING_KEY = "sort.str.descending";
		private final static String NUM_ASCENDING_KEY = "sort.num.ascending";
		private final static String NUM_DESCENDING_KEY = "sort.num.descending";

		/**
		 * Sort index has number
		 * <p> default false means sort target is string
		 */
		private boolean sortNumber;
		
		public SortAlgorithm(SortLevel sort) {
			this.sort = sort;
			setWidth("100%");
			setReadonly(true);
			setMold("rounded");
			appendComboitemsBySortIndex(sort.sortIndex);
			setSelectedIndex(sort.order == SortLevel.ASCENDING ? 0 : 1);
		}
		
		public void refresh() {
			setSelectedIndex(-1);
			Components.removeAllChildren(this);
			
			appendComboitemsBySortIndex(sort.sortIndex);
			setSelectedIndex(0);
		}
		
		private void appendComboitemsBySortIndex(int index) {
			if (index < 0) { //default sort target is string
				sortNumber = false;
				appendChild(new Comboitem(getLabel(STR_ASCENDING_KEY)));
				appendChild(new Comboitem(getLabel(STR_DESCENDING_KEY)));
			} else {
				if (isAllCellNumberType(index)) {
					sortNumber = true;
					appendChild(new Comboitem(getLabel(NUM_ASCENDING_KEY)));
					appendChild(new Comboitem(getLabel(NUM_DESCENDING_KEY)));
				} else {
					sortNumber = false;
					appendChild(new Comboitem(getLabel(STR_ASCENDING_KEY)));
					appendChild(new Comboitem(getLabel(STR_DESCENDING_KEY)));
				}
			}
		}
		
		/**
		 * Returns true indicate spreadsheet index column/row is number 
		 * @return
		 */
		private boolean isAllCellNumberType(int idx) {
			Worksheet sheet = ss.getSelectedSheet();
			Rect rect = ss.getSelection();
			int top = sortOrientation ? idx : rect.getTop();
			int left = sortOrientation ? rect.getLeft() : idx;
			int bottom = sortOrientation ? idx : rect.getBottom();
			int right = sortOrientation ? rect.getRight() : idx;

			for (int row = top; row <= bottom; row++) {
				for (int col = left; col <= right; col++) {
					Cell c = Utils.getCell(sheet, row, col);
					if (c != null) {
						int type = c.getCellType() != Cell.CELL_TYPE_FORMULA ?
									c.getCellType() : c.getCachedFormulaResultType();
						if (type != Cell.CELL_TYPE_NUMERIC)
							return false;
					}
				}
			}
			return true;
		}

		public void onSelect (SelectEvent evt) {
			Iterator<Comboitem> iter =  evt.getSelectedItems().iterator();
			if (iter.hasNext()) {
				Comboitem item = iter.next();
				if (!sortNumber) { //sort string
					sort.order = getLabel(STR_ASCENDING_KEY).equals(item.getLabel()) ?
							SortLevel.ASCENDING : SortLevel.DESCENDING;
				} else { //sort number
					sort.order = getLabel(NUM_ASCENDING_KEY).equals(item.getLabel()) ?
							SortLevel.ASCENDING : SortLevel.DESCENDING;
				}
			}
		}
	}
	
	/**
	 * Returns the string from i-18n, if not found, return empty string
	 * @return 
	 */
	private static String getLabel(String key) {
		String val = Labels.getLabel(key);
		return val != null ? val : "";
	}

	private static int getSpreadsheetIndexOffset (Spreadsheet spreadsheet, int index, boolean sortAlgorithm) {
		Rect rect = spreadsheet.getSelection();
		int baseIdx = sortAlgorithm == SORT_TOP_TO_BOTTOM ? rect.getLeft() : rect.getTop();
		int idx = index - baseIdx + 1;
		return idx >= 0 ? idx : -1;
	}
	
	
	/**
	 * Returns the index of column or row by title on selection range of spreadsheet
	 * @param spreadsheet
	 * @param title
	 * @param sortAlgorithm
	 * @return
	 */
	private static int getSpreadsheetIndexBy (Spreadsheet spreadsheet, String title, boolean sortAlgorithm) {
		Rect rect = spreadsheet.getSelection();
		if (sortAlgorithm == SORT_TOP_TO_BOTTOM) {
			for (int i = rect.getLeft(); i <= rect.getRight(); i++) {
				String t = spreadsheet.getColumntitle(i);
				if (Objects.equals(title, t))
					return i;
			}
		} else {
			for (int i = rect.getTop(); i <= rect.getBottom(); i++) {
				String t = spreadsheet.getRowtitle(i);
				if (Objects.equals(title, t))
					return i;
			}
		}
		return -1;
	}
}