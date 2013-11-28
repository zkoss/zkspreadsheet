package org.zkoss.zss.ngmodel;

public interface NPicture {

	public enum Format{
		PNG,
		JPG,
		GIF
	}
	public NSheet getSheet();
	
	public String getId();
	
	public Format getFormat();
	
	public byte[] getData();
	
	public NViewAnchor getAnchor();
}
