package zss.testapp.howto;

import java.util.Calendar;
import java.util.TimeZone;

import org.zkoss.poi.ss.usermodel.DateUtil;
import org.zkoss.util.TimeZones;

public class MyNow {

	public static double mynow(){
		Calendar sysnow = Calendar.getInstance();
		sysnow.add(Calendar.HOUR_OF_DAY, (TimeZones.getCurrent().getRawOffset()-TimeZone.getDefault().getRawOffset())/(1000*60*60));
		return DateUtil.getExcelDate(sysnow.getTime());
	}
}
