package org.zkoss.zss.ngmodel.sys.format;


public interface FormatEngine {

	FormatResult format(String format, Object value, FormatContext ctx);
	
}
