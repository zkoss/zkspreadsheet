/* NoCacheClassifier.java

	Purpose:
		
	Description:
		
	History:
		Apr 8, 2010 11:54:06 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.formula;

import org.apache.poi.hssf.record.formula.udf.UDFFinder;
import org.apache.poi.ss.formula.IStabilityClassifier;

/**
 * A classfier to deny caching of formula evaluation.
 * @author henrichen
 *
 */
public class NoCacheClassifier implements IStabilityClassifier {
	public static final IStabilityClassifier instance = new NoCacheClassifier();

	private NoCacheClassifier() {
		// enforce singleton
	}
	
	public boolean isCellFinal(int sheetIndex, int rowIndex, int columnIndex) {
		return true;
	}
}
