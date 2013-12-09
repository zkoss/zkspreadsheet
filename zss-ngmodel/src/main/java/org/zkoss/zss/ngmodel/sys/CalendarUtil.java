package org.zkoss.zss.ngmodel.sys;

import java.util.Date;

public interface CalendarUtil {

	public Date doubleValueToDate(double val, boolean date1904);
	public double dateToDoubleValue(Date value, boolean date1904);
}
