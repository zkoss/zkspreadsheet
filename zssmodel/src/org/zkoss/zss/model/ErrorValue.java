/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model;

import java.io.Serializable;
/**
 * An error result of a evaluated formula.
 * @author dennis
 * @since 3.5.0
 */
public class ErrorValue implements Serializable{
	private static final long serialVersionUID = 1L;
	/** <b>#NULL!</b>  - Intersection of two cell ranges is empty */
    public static final byte ERROR_NULL = 0x00;
    /** <b>#DIV/0!</b> - Division by zero */
    public static final byte ERROR_DIV_0 = 0x07;
    /** <b>#VALUE!</b> - Wrong type of operand */
    public static final byte INVALID_VALUE = 0x0F; 
    /** <b>#REF!</b> - Illegal or deleted cell reference */
    public static final byte ERROR_REF = 0x17;  
    /** <b>#NAME?</b> - Wrong function or range name */
    public static final byte INVALID_NAME = 0x1D; 
    /** <b>#NUM!</b> - Value range overflow */
    public static final byte ERROR_NUM = 0x24; 
    /** <b>#N/A</b> - Argument or function not available */
    public static final byte ERROR_NA = 0x2A;
    
    //TODO zss 3.5 this value is not in zpoi
    public static final byte INVALID_FORMULA = 0x7f;
	
	private byte _code;
	private String _message;

	public ErrorValue(byte code) {
		this(code, null);
	}

	public ErrorValue(byte code, String message) {
		this._code = code;
		this._message = message;
	}

	public byte getCode() {
		return _code;
	}

	public void setCode(byte code) {
		this._code = code;
	}

	public String getMessage() {
		return _message;
	}

	public void setMessage(String message) {
		this._message = message;
	}
	
	public String getErrorString(){
		return getErrorString(_code);
	}
	
	public String toString(){
		StringBuilder sb = new StringBuilder(super.toString());
		sb.append("/").append(getErrorString());
		return sb.toString();
	}
	
    public static final String getErrorString(int errorCode) {
        switch(errorCode) {
            case ERROR_NULL:  return "#NULL!";
            case ERROR_DIV_0: return "#DIV/0!";
            case INVALID_VALUE: return "#VALUE!";
            case ERROR_REF:   return "#REF!";
            case INVALID_NAME:  return "#NAME?";
            case ERROR_NUM:   return "#NUM!";
            case ERROR_NA:    return "#N/A";
            case INVALID_FORMULA:    return "#N/A";//TODO
        }
        return "#N/A";
//        throw new IllegalArgumentException("Bad error code (" + errorCode + ")");
    }

}
