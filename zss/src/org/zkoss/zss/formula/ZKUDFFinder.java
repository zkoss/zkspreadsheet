/* ZKUDFFinder.java

	Purpose:
		
	Description:
		
	History:
		Apr 6, 2010 4:54:57 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.formula;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.record.formula.functions.FreeRefFunction;
import org.apache.poi.hssf.record.formula.udf.UDFFinder;

/**
 * User defined function finder for ZK.
 * @author henrichen
 *
 */
public class ZKUDFFinder implements UDFFinder {
	private static Map<String, FreeRefFunction> FNS = new HashMap<String, FreeRefFunction>(3);
	static {
		FNS.put("ELEvaluate", ELEvaluate.instance);
	}
	public static final UDFFinder instance = new ZKUDFFinder();

	private ZKUDFFinder() {
		// enforce singleton
	}
	
	@Override
	public FreeRefFunction findFunction(String name) {
		return FNS.get(name);
	}
}
