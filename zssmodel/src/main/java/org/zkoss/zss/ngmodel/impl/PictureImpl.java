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

import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NViewAnchor;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class PictureImpl extends AbstractPictureAdv {

	String id;
	Format format;
	NViewAnchor anchor;
	byte[] data;
	AbstractSheetAdv sheet;

	public PictureImpl(AbstractSheetAdv sheet, String id, Format format,
			byte[] data, NViewAnchor anchor) {
		this.sheet = sheet;
		this.id = id;
		this.format = format;
		this.data = data;
		this.anchor = anchor;
	}
	@Override
	public NSheet getSheet(){
		checkOrphan();
		return sheet;
	}
	@Override
	public String getId() {
		return id;
	}
	@Override
	public Format getFormat() {
		return format;
	}
	@Override
	public NViewAnchor getAnchor() {
		return anchor;
	}

	@Override
	public void setAnchor(NViewAnchor anchor){
		this.anchor = anchor;
	}
	
	@Override
	public byte[] getData() {
		return data;
	}

	@Override
	public void destroy() {
		checkOrphan();
		sheet = null;
	}

	@Override
	public void checkOrphan() {
		if (sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
}
