/* MyFunction1.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jun 20, 2008 5:23:44 PM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.demo;

import java.util.Iterator;

/*import org.zkoss.zss.model.Cell;
import org.zkoss.zss.model.impl.RangeRef;
*/
import org.zkoss.poi.ss.usermodel.Cell;

/**
 * @author Dennis.Chen
 *
 */
public class CustomFunction {

	public static Object echo(java.lang.Object[] args, org.zkoss.xel.XelContext ctx){
//TODO associate poi evaluation to xel evaluation		
/*		StringBuffer sb = new StringBuffer();
		sb.append("[");
		if(args!=null){
			for(int i=0;i<args.length;i++){
				if(i>0) sb.append(", ");
				if(args[i] instanceof RangeRef){
					RangeRef ref = (RangeRef)args[i];
					for(Iterator iter = ref.getCells().iterator();iter.hasNext();){
						Cell cell = (Cell)iter.next();
						sb.append(cell.getText());
						if(iter.hasNext()) sb.append(", ");
					}
				}else if(args[i]!=null){
					sb.append(args[i].toString());
				}
			}
		}
		sb.append("]");
		return sb.toString();
*/		
		return "echo";
	}
}
