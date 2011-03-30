/* CircularIndex.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 29, 2011 4:42:46 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.zss.model.impl;

import java.util.Locale;

import org.zkoss.util.Locales;

/**
 * @author henrichen
 *
 */
public class CircularData {
	private final String[] _data;

	/*package*/ CircularData(String[] dataKey, int type) {
		_data = type == 1 ? getDataL(dataKey) : type == 2 ? getDataU(dataKey) : getDataN(dataKey);  
	}
	/*package*/ String getData(int current) {
		return _data[current];
	}
	/*package*/ int getIndex(String x) {
		for(int j = 0; j < _data.length; ++j) {
			if (_data[j].equalsIgnoreCase(x)) {
				return j;
			}
		}
		return -1;
	}
	/*package*/ int getSize() {
		return _data.length;
	}
	private String[] getDataN(String[] dataKey) { //normal. E.g. Sunday, Monday...
		//i18n
		return dataKey;
	}
	private String[] getDataL(String[] dataKey) { //all lowercase. E.g. sunday, monday...
		final String[] weekfn = getDataN(dataKey);
		final String[] weekL = new String[weekfn.length];
		final Locale current = Locales.getCurrent();
		for(int j = 0; j < weekfn.length; ++j) {
			weekL[j] = weekfn[j].toLowerCase(current);
		}
		return weekL;
	}
	private String[] getDataU(String[] dataKey) { //all uppercase. E.g. SUNDAY, MONDAY...
		final String[] weekfn = getDataN(dataKey);
		final String[] weekU = new String[weekfn.length];
		final Locale current = Locales.getCurrent();
		for(int j = 0; j < weekfn.length; ++j) {
			weekU[j] = weekfn[j].toUpperCase(current);
		}
		return weekU;
	}
}
