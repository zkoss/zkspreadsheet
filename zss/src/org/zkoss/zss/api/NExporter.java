package org.zkoss.zss.api;

import java.io.FileOutputStream;
import java.io.IOException;

import org.zkoss.zss.api.model.NBook;

public interface NExporter {
	
	public void export(NBook book, FileOutputStream fos) throws IOException;
}
