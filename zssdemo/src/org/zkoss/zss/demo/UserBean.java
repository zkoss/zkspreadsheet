/* UserBean.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 1, 2008 7:06:28 PM     2008, Created by Dennis.Chen
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
public class UserBean {

	private String firstName;
	private String lastName;
	
	public UserBean(String firstName,String lastName){
		this.firstName = firstName;
		this.lastName = lastName;
	}
	
	public String getFirstName() {
		return firstName;
	}
	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}
	public String getLastName() {
		return lastName;
	}
	public void setLastName(String lastName) {
		this.lastName = lastName;
	}
	
	public String getFullName(){
		return firstName+" "+lastName;
	}
	
}
