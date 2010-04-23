/* Report.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 2, 2008 5:52:50 PM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.demo;

/**
 * @author Dennis.Chen
 *
 */
public class Report {

	int val1;
	int val2;
	int val3;
	
	public Report(int val1,int val2,int val3){
		this.val1 = val1;
		this.val2 = val2;
		this.val3 = val3;
	}
		
	
	public int getVal1() {
		return val1;
	}
	public void setVal1(int val1) {
		this.val1 = val1;
	}
	public int getVal2() {
		return val2;
	}
	public void setVal2(int val2) {
		this.val2 = val2;
	}
	public int getVal3() {
		return val3;
	}
	public void setVal3(int val3) {
		this.val3 = val3;
	}
	
	
	
}
