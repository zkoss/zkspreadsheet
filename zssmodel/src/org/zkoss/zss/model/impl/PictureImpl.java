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

	private String _id;
	private Format _format;
	private ViewAnchor _anchor;
	private byte[] _data;
	private AbstractSheetAdv _sheet;

	public PictureImpl(AbstractSheetAdv sheet, String id, Format format,
			byte[] data, ViewAnchor anchor) {
		this._sheet = sheet;
		this._id = id;
		this._format = format;
		this._data = data;
		this._anchor = anchor;
	}
	@Override
	public SSheet getSheet(){
		checkOrphan();
		return _sheet;
	}
	@Override
	public String getId() {
		return _id;
	}
	@Override
	public Format getFormat() {
		return _format;
	}
	@Override
	public ViewAnchor getAnchor() {
		return _anchor;
	}

	@Override
	public void setAnchor(ViewAnchor anchor){
		this._anchor = anchor;
	}
	
	@Override
	public byte[] getData() {
		return _data;
	}

	@Override
	public void destroy() {
		checkOrphan();
		_sheet = null;
	}

	@Override
	public void checkOrphan() {
		if (_sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}
}
