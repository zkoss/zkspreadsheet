/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel.impl;

import java.util.Arrays;

import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class ObjectRefImpl extends RefImpl implements ObjectRef{

	private static final long serialVersionUID = 1L;
	
	private final String[] objectIdPath; 
	
	private final ObjectType objType;
	
	public ObjectRefImpl(AbstractChartAdv chart,String[] objectIdPath){
		super(RefType.OBJECT,chart.getSheet().getBook().getBookName(),chart.getSheet().getSheetName(), null,-1,-1,-1,-1);
		this.objectIdPath = objectIdPath;
		objType = ObjectType.CHART;
	}
	public ObjectRefImpl(AbstractChartAdv chart,String objectId){
		super(RefType.OBJECT,chart.getSheet().getBook().getBookName(),chart.getSheet().getSheetName(), null,-1,-1,-1,-1);
		this.objectIdPath = new String[]{objectId};
		objType = ObjectType.CHART;
	}

	@Override
	public String getObjectId() {
		return objectIdPath[objectIdPath.length-1];
	}

	@Override
	public String[] getObjectIdPath() {
		return objectIdPath;
	}
	@Override
	public ObjectType getObjectType() {
		return objType;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = super.hashCode();
		result = prime * result
				+ ((bookName == null) ? 0 : bookName.hashCode());
		result = prime * result
				+ ((sheetName == null) ? 0 : sheetName.hashCode());
		result = prime * result + ((objType == null) ? 0 : objType.hashCode());
		result = prime * result + Arrays.hashCode(objectIdPath);
		return result;
	}
	
	
	
	
	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (getClass() != obj.getClass())
			return false;
		ObjectRefImpl other = (ObjectRefImpl) obj;
		if (bookName == null) {
			if (other.bookName != null)
				return false;
		} else if (!bookName.equals(other.bookName))
			return false;
		if (sheetName == null) {
			if (other.sheetName != null)
				return false;
		} else if (!sheetName.equals(other.sheetName))
			return false;
		if (objType != other.objType)
			return false;
		if (!Arrays.equals(objectIdPath, other.objectIdPath))
			return false;

		return true;
	}
	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder(super.toString());
		sb.append(objType);
		for(String id:objectIdPath){
			sb.append(":").append(id);
		}
		return sb.toString();
	}
	
	
}
