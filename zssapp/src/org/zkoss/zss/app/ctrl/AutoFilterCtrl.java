/* AutoFilterCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		May 13, 2011 10:24:13 AM , Created by sam
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.ctrl;

import java.util.Calendar;
import java.util.Comparator;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import java.util.TreeSet;

import org.zkoss.lang.Objects;
import org.zkoss.lang.Strings;
//import org.zkoss.poi.ss.usermodel.FilterColumn;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.AutoFilterOperation;
//import org.zkoss.zss.api.Range.CellType;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zul.Button;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.ListitemRenderer;

/**
 * @author sam
 *
 */
public class AutoFilterCtrl extends GenericForwardComposer {
	
	private Dialog _autoFilterDialog;
	private Listbox _filterListbox;
	
	private Button okBtn;
	
	private final static String KEY_ROW_INFO = "rowInfo";
	private final static String BLANK_DISPLAY = "(Blanks)";
	private final static Comparable BLANK_VALUE = new Comparable() {
		@Override
		public int compareTo(Object o) {
			return BLANK_VALUE.equals(o) ? 0 : 1; //unless same otherwise BLANK_VALUE is always the biggest!
		}
	}; 
	private final static RowInfo BLANK_ROW_INFO = new RowInfo(BLANK_VALUE, BLANK_DISPLAY);

	private int fieldIndex = 0;
	private int columnIndex = 0;
	private Sheet worksheet;
	private Range range;
	
	private boolean isHiddenRow(Sheet sheet, int rowIdx) {
		return sheet.isRowHidden(rowIdx);
	}
	
	private void fetchRowInfos(FilterColumn fc, Range range, Set<RowInfo> all, Set<RowInfo> selected) {
		final Set criteria1 = fc == null ? null : fc.getCriteria1();
		final boolean nofilter = criteria1 == null || criteria1.isEmpty(); 
		boolean hasBlank = false;
		boolean selectedBlank = false;
		final int top = range.getRow() + 1;
		final int bottom = range.getLastRow();
		for (int i = top; i <= bottom; i++) {
			if (nofilter && isHiddenRow(worksheet, i)) {
				continue;
			}
			Range c = Ranges.range(worksheet,i,columnIndex);
			final boolean blankcell = c.getCellData().getType()==CellType.BLANK;
			if (!blankcell) {
				String displaytxt = c.getCellData().getFormatText();
				Object val = c.getCellValue();
				
				RowInfo rowInfo = new RowInfo(val, displaytxt); 
				all.add(rowInfo);
				if (criteria1 == null || criteria1.isEmpty() || criteria1.contains(displaytxt)) { //selected
					selected.add(rowInfo);
				}
			} else {
				hasBlank = true;
				if (!selectedBlank && (nofilter || criteria1.contains("="))) { //selected
					selectedBlank = true;
				}
			}
		}
		if (hasBlank) {
			all.add(BLANK_ROW_INFO);
		}
		if (selectedBlank) {
			selected.add(BLANK_ROW_INFO);
		}
	}
	
	private static class MyComparator implements Comparator<RowInfo> {		
		@Override
		public int compare(RowInfo o1, RowInfo o2) {
			final Object val1 = o1.value;
			final Object val2 = o2.value;
			final int type1 = getType(val1);
			final int type2 = getType(val2);
			final int typediff = type1 - type2;
			if (typediff != 0) {
				return typediff;
			}
			switch(type1) {
			case 1: //Date
				return compareDates((Date)val1, (Date)val2);
			case 2: //Number
				return ((Double)val1).compareTo((Double)val2);
			case 3: //String
				return ((String)val1).compareTo((String)val2);
			case 4: //Boolean
				final boolean b1 = ((Boolean)val1).booleanValue();
				final boolean b2 = ((Boolean)val2).booleanValue();
				return !b1 && b2 ? -1 : b1 && !b2 ? 1 : 0;
			case 5: //Error(Byte)
				return ((Byte)val1).compareTo((Byte)val2);
			default:
			case 6: //(Blanks)
				return 0;
			}
		}
		private int compareDates(Date val1, Date val2) {
			final Calendar cal1 = Calendar.getInstance();
			final Calendar cal2 = Calendar.getInstance();
			cal1.setTime((Date)val1);
			cal2.setTime((Date)val2);
			
			//year
			final int y1 = cal1.get(Calendar.YEAR);
			final int y2 = cal2.get(Calendar.YEAR);
			final int ydiff = y2 - y1; //bigger year is less in sorting
			if (ydiff != 0) {
				return ydiff;
			}
			
			//month
			final int m1 = cal1.get(Calendar.MONTH);
			final int m2 = cal2.get(Calendar.MONTH);
			final int mdiff = m1 - m2; 
			if (mdiff != 0) {
				return mdiff;
			}
			
			//day
			final int d1 = cal1.get(Calendar.DAY_OF_MONTH);
			final int d2 = cal2.get(Calendar.DAY_OF_MONTH);
			final int ddiff = d1 - d2; //smaller month is bigger in sorting 
			if (ddiff != 0) {
				return ddiff;
			}
			
			//hour
			final int h1 = cal1.get(Calendar.HOUR_OF_DAY);
			final int h2 = cal2.get(Calendar.HOUR_OF_DAY);
			final int hdiff = h1 - h2;
			if (hdiff != 0) {
				return hdiff;
			}
			
			//minutes
			final int mm1 = cal1.get(Calendar.MINUTE);
			final int mm2 = cal2.get(Calendar.MINUTE);
			final int mmdiff = mm1 - mm2;
			if (mmdiff != 0) {
				return mmdiff;
			}
			
			//seconds
			final int s1 = cal1.get(Calendar.SECOND);
			final int s2 = cal2.get(Calendar.SECOND);
			final int sdiff = s1 - s2;
			if (sdiff != 0) {
				return sdiff;
			}
			
			//millseconds
			final int ms1 = cal1.get(Calendar.MILLISECOND);
			final int ms2 = cal2.get(Calendar.MILLISECOND);
			return ms1 - ms2;
		}
		//Date < Number < String < Boolean(FALSE < TRUE) < Error(byte) < (Blanks)
		private int getType(Object val) {
			if (val instanceof Date) {
				return 1;
			}
			if (val instanceof Byte) { //error
				return 5;
			}
			if (val instanceof Number) {
				return 2;
			}
			if (val instanceof String) {
				return Strings.isEmpty((String)val) ? 6 : 3;
			}
			if (val instanceof Boolean) {
				return 4;
			}
			return 6;
		}
	}
	public void onOpen$_autoFilterDialog(ForwardEvent evt) {
		Object[] info = (Object[]) evt.getOrigin().getData();
		range = (Range) info[2];
		worksheet = (Sheet) range.getSheet(); 
		if (!worksheet.isAutoFilterEnabled()) {
			return;
		}
		fieldIndex = (Integer) info[0];
		columnIndex = (Integer) info[1];
		
		
		//TODO, Dennis , need to wrap to model
		final FilterColumn fc = ((SheetImpl)worksheet).getNative().getAutoFilter().getFilterColumn(fieldIndex - 1);
		
		final TreeSet<RowInfo> rowInfos = new TreeSet<RowInfo>(new MyComparator());
		final Set<RowInfo> selected = new HashSet<RowInfo>();
		fetchRowInfos(fc, range, rowInfos, selected);
		
		_filterListbox.setModel(new ListModelList(rowInfos));
		
		//handle selection
		ListModelList model = (ListModelList) _filterListbox.getListModel();
		for(RowInfo rowInfo : selected) {
			model.addSelection(rowInfo);
		}
		
		_filterListbox.setItemRenderer(new ListitemRenderer() {
			public void render(Listitem item, Object data) throws Exception {
				RowInfo info = (RowInfo)data;
				item.setLabel(info.display);
			}

			@Override
			public void render(Listitem item, Object data, int index)
					throws Exception {
				render(item, data);
			}
		});
	}
	
	private static class RowInfo {
		private Object value;
		private String display;
		
		RowInfo(Object val, String displayVal) {
			value = val;
			display = displayVal;
		}

		public int hashCode() {
			return value == null ? 0 : value.hashCode();
		}

		@Override
		public boolean equals(Object obj) {
			if (this == obj)
				return true;
			if (!(obj instanceof RowInfo))
				return false;
			final RowInfo other = (RowInfo) obj;
			return Objects.equals(other.value, this.value);
		}
	}

	public void onClick$okBtn() {
		final ListModelList model = (ListModelList) _filterListbox.getListModel();
		final int itemcount = model.size();
		final Set selected = model.getSelection();
		final int selcount = selected.size();
		if (selcount < itemcount) { //partial selection
			final String[] criteria = new String[selcount];
			int j = 0;
			final TreeSet<RowInfo> selset = new TreeSet<RowInfo>(new MyComparator());
			selset.addAll(selected);
			for (final Iterator<RowInfo> it = selset.iterator(); it.hasNext(); ) {
				RowInfo info = (RowInfo) it.next();
				criteria[j++] = BLANK_ROW_INFO.equals(info) ? "=" : info.display;
			}
			range.enableAutoFilter(fieldIndex, AutoFilterOperation.VALUES, criteria , null, null);
		} else { //select all!
			range.enableAutoFilter(fieldIndex, AutoFilterOperation.VALUES, null , null, null);
		}
		_autoFilterDialog.fireOnClose(null);
	}
}