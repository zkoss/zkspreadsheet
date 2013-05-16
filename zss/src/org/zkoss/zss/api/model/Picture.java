package org.zkoss.zss.api.model;


public interface Picture {
	
	public enum Format{
	    EMF,
	    WMF,
	    PICT,
	    JPEG,
	    PNG,
	    DIB		
	}
	
	public String getId();
}
