/* MergedRect.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 19, 2008 2:54:10 PM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import org.zkoss.zss.api.Rect;

/**
 * @author Dennis.Chen
 *
 */
public class MergedRect extends Rect{

	private int _id;
	public MergedRect(){
		this(-1,-1,-1,-1,-1);
	}
	public MergedRect(int id,int left,int top,int right,int bottom){
		super(left,top,right,bottom);
		this._id = id;
	}
	
	public int getId(){
		return _id;
	}
	
	public void setId(int id){
		_id = id;
	}

	@Override
	public Object cloneSelf(){
		return new MergedRect(_id,getLeft(),getTop(),getRight(),getBottom());
	}
	
	@Override
	public String toString(){
		return "id:"+getId()+",left:"+getLeft()+",top:"+getTop()+",right:"+getRight()+",bottom:"+getBottom();
	}
	
	@Override
	public int hashCode() {
		return _id ^ super.hashCode();
	}
	
	@Override
	public boolean equals(Object other) {
		return this == other
			|| (other instanceof MergedRect
					&& _id == ((MergedRect)other)._id);
	}
}
