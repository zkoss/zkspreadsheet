package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NViewAnchor;

public class PictureImpl extends AbstractPicture {

	String id;
	Format format;
	NViewAnchor anchor;
	byte[] data;
	AbstractSheet sheet;
	
	
	public PictureImpl(AbstractSheet sheet,String id, Format format, byte[] data,NViewAnchor anchor){
		this.sheet = sheet;
		this.id = id;
		this.format = format;
		this.data = data;
		this.anchor = anchor;
	}
	public AbstractSheet getSheet(){
		checkOrphan();
		return sheet;
	}
	public String getId() {
		return id;
	}

	public Format getFormat() {
		return format;
	}
	public NViewAnchor getAnchor() {
		return anchor;
	}
	public void setAnchor(NViewAnchor anchor) {
		this.anchor = anchor;
	}
	public byte[] getData() {
		return data;
	}
	@Override
	public void release() {
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
