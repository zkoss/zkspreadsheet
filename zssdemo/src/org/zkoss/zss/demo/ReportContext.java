/* ReportContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 2, 2008 5:52:41 PM     2008, Created by Dennis.Chen
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
public class ReportContext {

	Report report = new Report(0,0,0);

	public Report getReport() {
		return report;
	}

	public void setReport(Report report) {
		this.report = report;
	}
	
	
	
}
