package org.zkoss.zss.ngmodel.sys.format;

import org.zkoss.zss.ngmodel.NCell;

public interface FormatEngine {

	FormatResult format(String format, Object value, FormatContext ctx);
	
}
