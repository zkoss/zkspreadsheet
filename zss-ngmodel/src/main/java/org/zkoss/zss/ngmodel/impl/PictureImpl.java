package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NViewAnchor;

public class PictureImpl extends AbstractPicture {

	String id;
	Format format;
	NViewAnchor anchor;
	byte[] data;
	
	
	public PictureImpl(String id, Format format, byte[] data,NViewAnchor anchor){
		this.id = id;
		this.format = format;
		this.data = data;
		this.anchor = anchor;
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
}
