package org.zkoss.zss.ngmodel.format;

import org.zkoss.zss.ngmodel.NCell;

public interface FormatEngine {

	FormatResult format(NCell cell, FormatContext ctx);
	
}
