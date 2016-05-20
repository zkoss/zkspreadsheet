/* FilterRowInfoComparator.java

	Purpose:
		
	Description:
		
	History:
		May 19, 2016 5:06:46 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.util.Calendar;
import java.util.Comparator;
import java.util.Date;

import org.zkoss.lang.Strings;
import org.zkoss.zss.model.ErrorValue;

/**
 * @author henri
 * @since 3.9.0
 */
public class FilterRowInfoComparator implements Comparator<FilterRowInfo> {		
	@Override
	public int compare(FilterRowInfo o1, FilterRowInfo o2) {
		final Object val1 = o1.getValue();
		final Object val2 = o2.getValue();
		final int type1 = FilterRowInfo.getType(val1);
		final int type2 = FilterRowInfo.getType(val2);
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
			//ZSS-935
			final byte by1 = val1 instanceof ErrorValue ? Byte.valueOf(((ErrorValue)val1).getCode()) : ((Byte) val1).byteValue();
			final byte by2 = val2 instanceof ErrorValue ? Byte.valueOf(((ErrorValue)val2).getCode()) : ((Byte) val2).byteValue();
			return by1 - by2;
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
}