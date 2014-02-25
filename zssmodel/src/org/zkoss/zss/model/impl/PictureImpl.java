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

import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.ViewAnchor;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class PictureImpl extends AbstractPictureAdv {

	String id;
	Format format;
	ViewAnchor anchor;
	byte[] data;
	AbstractSheetAdv sheet;

	public PictureImpl(AbstractSheetAdv sheet, String id, Format format,
			byte[] data, ViewAnchor anchor) {
		this.sheet = sheet;
		this.id = id;
		this.format = format;
		this.data = data;
		this.anchor = anchor;
	}
	@Override
	public SSheet getSheet(){
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
	public ViewAnchor getAnchor() {
		return anchor;
	}

	@Override
	public void setAnchor(ViewAnchor anchor){
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
