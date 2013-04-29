package org.zkoss.zss.api;

import java.io.IOException;
import java.io.InputStream;

import org.zkoss.zss.api.model.NBook;

public interface NImporter {
		
	public NBook imports(InputStream is, String bookName) throws IOException;
	
}
