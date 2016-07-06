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
package org.zkoss.zss.model.impl;

import java.util.Set;

import org.zkoss.zss.model.SConditionalFormatting;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.sys.dependency.ConditionalRef;
import org.zkoss.zss.model.sys.dependency.Ref;
/**
 * 
 * @author henrichen
 * @since 3.9.0
 */
public class ConditionalRefImpl extends RefImpl implements ConditionalRef{

	private static final long serialVersionUID = 1L;
	
	private final int _conditionalId;
	
	public ConditionalRefImpl(SConditionalFormatting cfmt){
		this(cfmt.getSheet().getBook().getBookName(),cfmt.getSheet().getSheetName(),cfmt.getId());
	}
	public ConditionalRefImpl(String bookName, String sheetName, int id) {
		super(RefType.CONDITIONAL,bookName,sheetName, null, -1,-1,-1,-1);
		this._conditionalId = id;
	}

	
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = super.hashCode();
		result = prime * result
				+ ((bookName == null) ? 0 : bookName.hashCode());
		result = prime * result
				+ ((sheetName == null) ? 0 : sheetName.hashCode());
		result = prime * result	+ _conditionalId;
		return result;
	}
	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (getClass() != obj.getClass())
			return false;
		ConditionalRefImpl other = (ConditionalRefImpl) obj;
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
		return _conditionalId == other._conditionalId;
	}
	
	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder(super.toString());
		if(sheetName!=null){
			sb.append(sheetName).append(":");
		}
		sb.append(_conditionalId);
		return sb.toString();
	}
	@Override
	public int getConditionalId() {
		return _conditionalId;
	}
}
