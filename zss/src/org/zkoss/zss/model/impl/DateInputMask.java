/* DateInputFormat.java

	Purpose:
		
	Description:
		
	History:
		Oct 1, 2010 10:05:57 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.Calendar;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Responsible for date input mask. 
 * @author henrichen
 *
 */
public class DateInputMask {
	private static final String SPACES = "[ \\t]*";
	private static final String SPACE1 = "[ \\t]+";
	private static final String START_ANCHOR = "^"+SPACES;
	private static final String END_ANCHOR = SPACES+"$";
	private static final String AMPM = "[AP]M?";
	private static final String AMPM1_G = "("+AMPM+")";
	private static final String DASH_SLASH_SPACE = "(?:"+SPACES+"[-/]?|[-/]?)"+SPACES; // dash, slash, or space
	private static final String DOT_SPACE = "(?:"+SPACES+"[.]?|[.]?)"+SPACES; //dot or space
	private static final String COLON_SPACE = "(?:"+SPACES+"[:]?|[:]?)"+SPACES; // colon or space
	private static final String DASH_SLASH = SPACES+"[/-]"+SPACES; // dash or slash
	private static final String COLON_AMPM = "(?:"+SPACES+"[:]|"+SPACE1+AMPM1_G+")"+SPACES; //colon or am/pm 
	private static final String COMMA = SPACES+"[,]"+ SPACE1;
	private static final String NAME_MONTH_G = "([A-Z]{3,9})"; //Named Month: Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec
	private static final String NUM_MONTH_G = "([0]?[1-9]|1[012])"; //Numbed Month: 1 ~ 12
	private static final String DAY_G = "([0]?[1-9]|[12][0-9]|3[01])"; //Day: 1 ~ 31
	private static final String YEAR_G = "(19\\d\\d|[2-9]\\d\\d\\d|\\d\\d|\\d)"; //1900 - 9999
	
	private static final String HOUR_G = "(\\d{1,4})";
	private static final String MINUTE_G = "(\\d{1,4})";
	private static final String SECOND_G = "(\\d{0,4})";
	private static final String MSECOND_G = "(\\d{0,4})";
	private static final String AMPM2_G = "("+SPACE1+AMPM+"|"+AMPM+")?";
	private static final String TIME = HOUR_G + COLON_AMPM 
		+ "(?:" + MINUTE_G + COLON_SPACE + "|" + SPACES + ")" 
		+ SECOND_G + DOT_SPACE + MSECOND_G + AMPM2_G;
	private static final String TIME_PAT = START_ANCHOR + TIME + END_ANCHOR;
	
	private static final String DATE11 = NAME_MONTH_G + DASH_SLASH_SPACE + DAY_G + COMMA + YEAR_G; //Feb 1, 2010 || Feb-1, 2010 => apply "d-mmm-yy"   
	private static final String DATE12 = DAY_G + DASH_SLASH_SPACE + NAME_MONTH_G + DASH_SLASH_SPACE + YEAR_G; //1 Feb 2010 || 1 - Feb - 2010 || 1 / Feb / 2000 => apply "d-mmm-yy"  
	private static final String DATE21 = NAME_MONTH_G + DASH_SLASH_SPACE + DAY_G; //Feb 1 || Feb-1 || Feb/1 => apply "d-mmm"  
	private static final String DATE22 = NUM_MONTH_G + DASH_SLASH + DAY_G; //2-1 || 2/1  => apply "d-mmm", might apply "mmm-yy" if date over month's limit  
	private static final String DATE31 = NAME_MONTH_G + DASH_SLASH_SPACE + YEAR_G; //Feb 31 || Feb-31 || Feb/31 => apply "mmm-yy"
	private static final String DATE32 = NUM_MONTH_G + DASH_SLASH + YEAR_G; //2-31 || 2/31 => apply "mmm-yy"
	private static final String DATE4 = NUM_MONTH_G + DASH_SLASH + DAY_G + DASH_SLASH + YEAR_G; //2-1-2010 || 2/1/2010 || 2/1-2010 || 2-1/2010 => apply "m/d/yyyy"

	private static final String DATE11_PAT = START_ANCHOR + DATE11 + END_ANCHOR;   
	private static final String DATE12_PAT = START_ANCHOR + DATE12 + END_ANCHOR;  
	private static final String DATE21_PAT = START_ANCHOR + DATE21 + END_ANCHOR;  
	private static final String DATE22_PAT = START_ANCHOR + DATE22 + END_ANCHOR;  
	private static final String DATE31_PAT = START_ANCHOR + DATE31 + END_ANCHOR;
	private static final String DATE32_PAT = START_ANCHOR + DATE32 + END_ANCHOR;
	private static final String DATE4_PAT = START_ANCHOR + DATE4 + END_ANCHOR;
	
	private static final String DATETIME11_PAT = START_ANCHOR + DATE11 + SPACE1 + TIME + END_ANCHOR;   
	private static final String DATETIME12_PAT = START_ANCHOR + DATE12 + SPACE1 + TIME + END_ANCHOR;  
	private static final String DATETIME21_PAT = START_ANCHOR + DATE21 + SPACE1 + TIME + END_ANCHOR;  
	private static final String DATETIME22_PAT = START_ANCHOR + DATE22 + SPACE1 + TIME + END_ANCHOR;  
	private static final String DATETIME31_PAT = START_ANCHOR + DATE31 + SPACE1 + TIME + END_ANCHOR;
	private static final String DATETIME32_PAT = START_ANCHOR + DATE32 + SPACE1 + TIME + END_ANCHOR;
	private static final String DATETIME4_PAT = START_ANCHOR + DATE4 + SPACE1 + TIME + END_ANCHOR;

	private static final FormatInfo[] _formatInfo= new FormatInfo[15];
	
	static {
        final int flags = Pattern.CASE_INSENSITIVE;
        int j = 0;
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATE11_PAT, flags), "d-mmm-yy", "JUL 18, 65", null); //0
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATE12_PAT, flags), "d-mmm-yy", "18 JUL 65", null); //1
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATE21_PAT, flags), "d-mmm", "JUL 18", null); //2
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATE22_PAT, flags), "d-mmm", "7-18", null); //3
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATE31_PAT, flags), "mmm-yy", "JUL 65", null); //4
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATE32_PAT, flags), "mmm-yy", "7-65", null); //5
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATE4_PAT, flags), "m/d/yyyy", "7-18-65", null); //6
        _formatInfo[j++] = new FormatInfo(Pattern.compile(TIME_PAT, flags), null, "2:3:5.100 A", "2 P"); //7
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATETIME11_PAT, flags), "m/d/yyyy h:mm", "JUL 18, 65 2:3:5.100 A", "JUL 18, 65 2 P"); //8 
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATETIME12_PAT, flags), "m/d/yyyy h:mm", "18 JUL 65 2:3:5.100 A", "18 JUL 65 2 P");  //9
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATETIME21_PAT, flags), "m/d/yyyy h:mm", "JUL 18 2:3:5.100 A", "JUL 18 2 P"); //10
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATETIME22_PAT, flags), "m/d/yyyy h:mm", "7-18 2:3:5.100 A", "7-18 2 P"); //11
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATETIME31_PAT, flags), "m/d/yyyy h:mm", "JUL 65 2:3:5.100 A", "JUL 65 2 P"); //12
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATETIME32_PAT, flags), "m/d/yyyy h:mm", "7-65 2:3:5.100 A", "7-65 2 P"); //13
        _formatInfo[j++] = new FormatInfo(Pattern.compile(DATETIME4_PAT, flags), "m/d/yyyy h:mm", "7-18-65 2:3:5.100 A", "7-18-65 2 P"); //14
	}
	
	public Object[] parseDateInput(String txt) {
		for(int j = 0; j < _formatInfo.length; ++j) {
			final FormatInfo info = _formatInfo[j]; 
	        Matcher m = info.getMaskPattern().matcher(txt);
	        if (m.matches()) {
	        	return info.parseInput(txt, m); 
	        }
		}
		return new Object[] {txt, null};
	}
	private static class FormatInfo {
		private final Pattern _mask;
		private final String _format;
		private int _year;
		private int _month;
		private int _day;
		private int _hour;
		private int _minute;
		private int _second;
		private int _msecond;
		private int _ampm1;
		private int _ampm2;
		
		//match text: "Jul 18, 1965 2:3:5.100 A"
		public FormatInfo(Pattern mask, String format, String groupMatchText, String ampm1MatchText) {
			_mask = mask;
			_format = format;
	        Matcher m = _mask.matcher(groupMatchText);
	        if (!m.matches()) {
	        	throw new RuntimeException("Wrong groupMatchText: "+ groupMatchText+ ", regex mask: "+_mask);
	        }
	        for (int j = 1, len = m.groupCount()+1; j < len; ++j) {
	        	final String grptxt = m.group(j);
	        	if ("65".equals(grptxt)) {
	        		_year = j;
	        	} else if ("7".equals(grptxt)) {
	        		_month = j;
	        	} else if ("JUL".equals(grptxt)) {
	        		_month = j;
	        	} else if ("18".equals(grptxt)) {
	        		_day = j;
	        	} else if (" A".equals(grptxt)) {
	        		_ampm2 = j;
	        	} else if ("2".equals(grptxt)) {
	        		_hour = j;
	        	} else if ("3".equals(grptxt)) {
	        		_minute = j;
	        	} else if ("5".equals(grptxt)) {
	        		_second = j;
	        	} else if ("100".equals(grptxt)) {
	        		_msecond = j;
	        	}
	        }
	        if (ampm1MatchText != null) {
		        m = _mask.matcher(ampm1MatchText);
		        if (!m.matches()) {
		        	throw new RuntimeException("Wrong ampm1MatchText: "+ ampm1MatchText+ ", regex mask: "+_mask);
		        }
		        for (int j = 1, len = m.groupCount(); j < len; ++j) {
		        	final String grptxt = m.group(j);
		        	if ("P".equals(grptxt)) {
		        		_ampm1 = j;
		        		break;
		        	}
		        }
	        }
		}
		
		public Pattern getMaskPattern() {
			return _mask;
		}
		
		public String getFormat() {
			return _format;
		}
		
		private static final String[] MONTHS = {
			"JANUARY",
			"FEBRUARY",
			"MARCH",
			"APRIL",
			"MAY",
			"JUNE",
			"JULY",
			"AUGUST",
			"SEPTEMBER",
			"OCTOBER",
			"NOVEMBER",
			"DECEMBER"
		};
			
		private int getMonthIndex(String month) {
			month = month.toUpperCase();
			for(int j = 0; j < MONTHS.length; ++j) {
				if (MONTHS[j].startsWith(month))
					return j;
			}
			return -1;
		}
		
		//return month index (January is 0)
		private int parseMonth(Matcher m) {
			if (_month > 0) {
				final String month = m.group(_month);
				if (month.length() >= 3) { //Feb case
					return getMonthIndex(month);
				} else {
					return Integer.parseInt(month) - 1;
				}
			}
			return -1;
		}
		
		//return day index (1st day is 1)
		private int parseDay(Matcher m) {
			if (_day > 0) {
				final String day = m.group(_day);
				return Integer.parseInt(day);
			}
			return -1;
		}
		
		//return year
		private int parseYear(Matcher m) {
			if (_year > 0) {
				final String year = m.group(_year);
				int y = Integer.parseInt(year);
				return y < 100 ? (2000+y) : y;
			}
			return -1;
		}
		private static final int[] MAXDAYS = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
		//given month and year, return the maxday
		private int getMaxday(int month, int year) {
			int maxday = MAXDAYS[month];
			if (month == 2 && isLeapYear(year)) {
				return maxday + 1;
			}
			return maxday;
		}
		private boolean isLeapYear(int year) {
			return year % 4 == 0 && (year % 100 != 0 || year % 400 == 0); 
		}
		//return hour number
		private int parseHour(Matcher m) {
			return parseTime(m, _hour);
		}
		//return minute number
		private int parseMinute(Matcher m) {
			return parseTime(m, _minute);
		}
		//return second number
		private int parseSecond(Matcher m) {
			return parseTime(m, _second);
		}
		//return millsecnd number
		private int parseMillisecond(Matcher m) {
			return parseTime(m, _msecond);
		}
		//return minute number
		private int parseTime(Matcher m, int groupIndex) {
			if (groupIndex > 0) {
				final String value = m.group(groupIndex);
				if (value != null && value.length() > 0) {
					return Integer.parseInt(value);
				}
			}
			return -1;
		}
		//return ampm
		private String parseAmpm(Matcher m) {
			String result = null; 
			if (_ampm2 > 0) {
				result = m.group(_ampm2);
			} 
			if (result == null && _ampm1 > 0) {
				result = m.group(_ampm1);
			}
			return result;
		}
		
		//TODO locale and timezone for the calendar?
		private Calendar getCalendar() {
			return Calendar.getInstance();
		}
		public Object[] parseInput(String txt, Matcher m) {
			int month = parseMonth(m);
			if (month < 0 && _month > 0) {
				return new Object[] {txt, null};
			}
			int day = parseDay(m);
			if (day < 0 && _day > 0) {
				return new Object[] {txt, null};
			}
			final Calendar cal = getCalendar();
			int year = parseYear(m);
			if (year < 0 && _format != null) {
				year = cal.get(Calendar.YEAR); 
			}
			String format = _format;
			if (_month > 0) {
				int maxday = getMaxday(month, year);
				if (maxday < day) { // illegal date
					if (_year > 0) {  
						return new Object[] {txt, null};
					} else { //month-year case
						if (_hour <= 0) { //if no hour
							format = "mmm-yy";
						}
						year = day + 1900;
						day = 1;
					}
				}
			}
			int hour = parseHour(m);
			int minute = parseMinute(m);
			int second = parseSecond(m);
			int msecond = parseMillisecond(m);
			
			String ampm = parseAmpm(m);
			
			if (ampm != null) {
				//invalid hour, minute, second with am/pm
				if (hour > 12 || minute >= 60 || second >= 60) {
					return new Object[] {txt, null};
				}
				ampm = ampm.toUpperCase();
				if (hour < 12) {
					if (ampm.startsWith("P")) {
						hour += 12;
					}
				} else { //hour == 12
					if (ampm.startsWith("A")) {
						hour = 0;
					}
				}
			}
			
			if ((hour > 23 && minute >= 60)
				|| (minute >= 60 && second >=60)
				|| (hour > 23 && second >= 60)) {
				return new Object[] {txt, null};
			}
			
			if (format == null) { //pure time pattern
				year = 1900;
				month = 1 - 1; //0 based index
				day = 1;
				if (msecond >= 0) {
					format = "mm:ss.0";
				} else if (ampm != null) {
					format = second >= 0 ? "h:mm:ss AM/PM" : "h:mm AM/PM";
				} else if (hour < 24) {
					if ( minute < 60 && second < 60) {
						format = second > 0 ? "h:mm:ss" : "h:mm";
					} // else a number; so format == null
				} else {
					if (minute < 60 && second < 60) {
						format = "[h]:mm:ss";
					} // else a number; so format == null
				}
			} else {
				if (msecond >= 0) {
					format = "mm:ss.0";
				} else if (hour > 23 || minute >= 60 || second >=60) {
					format = null; //a number
				}
			}
			
			//date
			cal.set(Calendar.YEAR, year);
			cal.set(Calendar.MONTH, month);
			cal.set(Calendar.DAY_OF_MONTH, day);
			
			//time
			if (hour < 0) {
				hour = 0;
			}
			if (minute < 0) {
				minute = 0;
			}
			if (second < 0) {
				second = 0;
			}
			if (msecond < 0) {
				msecond = 0;
			}
			cal.set(Calendar.HOUR_OF_DAY, hour);
			cal.set(Calendar.MINUTE, minute);
			cal.set(Calendar.SECOND, second);
			
			cal.set(Calendar.MILLISECOND, msecond);
			return new Object[] {cal.getTime(), format}; 
		}
	}
}
