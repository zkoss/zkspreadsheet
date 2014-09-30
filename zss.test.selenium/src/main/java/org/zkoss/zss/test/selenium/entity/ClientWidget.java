/* Script.java

	Purpose:
		
	Description:
		
	History:
		Fri, Jul 11, 2014  5:05:17 PM, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.test.selenium.entity;

import java.util.LinkedList;
import java.util.List;

import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.TestCaseBase;

/**
 * The skeleton of ZK client side widget. It is used to manipulate a string buffer
 * to concatenate an executed JavaScript code.
 * @author jumperchen
 * @author RaymondChao
 */
public class ClientWidget {
	
	private String selector;
	private List<String> js;
	
	private static String FUNC = "%1$s()";
	private static String FUNC_ARG = "%1$s(%2$s)";
	protected static String RETURN_FUN = "return %1$s.%2$s();";
	private static String GET = "return %1$s.get%2$s();";
	
	protected ClientWidget() {
		js = new LinkedList<String>();
	}
	/**
	 * Returns true if the string is null or empty.
	 */
	public static final boolean isEmpty(String string) {
		return string == null || string.length() == 0;
	}
	
	public void setSelector(String selector) {
		this.selector = selector; 
	}
	
	public String getSelector() {
		return selector;
	}
	
	public void add(String snippet) {
		js.add(snippet);
	}
	
	/**
	 * Calls a JavaScript Function without any arguments.
	 * @param funcName the name of the JavaScript function without parentheses.
	 */
	public ClientWidget call(String funcName) {
		add(String.format(FUNC, funcName));
		return this;
	}

	/**
	 * Calls a JavaScript Function with any arguments.
	 * @param funcName the name of the JavaScript function without parentheses.
	 * @param arguments the arguments in ordering.
	 */
	public ClientWidget call(String funcName, Object... arguments) {
		StringBuilder args = new StringBuilder();
		for (Object arg : arguments) {
			args.append(arg).append(',');
		}
		if (arguments.length > 0) {
			args.deleteCharAt(args.length() - 1);
		}

		add(String.format(FUNC_ARG, funcName, args.toString()));
		return this;
	}
	
	/**
	 * Returns the string that the first word of the name is upper case.
	 * @param key the prefix of the method name. Like <code>set</code> and <code>get</code>
	 * @param name the name of the method.
	 */
	protected String toUpperCase(String key, String name) {
		char[] buf = name.toCharArray();
		buf[0] = Character.toUpperCase(buf[0]);
		return key + new String(buf);
	}
	
	public boolean isEmpty() {
		return js.isEmpty();
	}
	
	/**
	 * Returns the string in a return value of script.
	 * <p> For example,
	 * <pre><code>return YourCodesHere;</code></pre>
	 */
	public String getResult(String fun) {
		return String.format(RETURN_FUN, toString(), fun);
	}
	
	public Object get(String fun) {
		return TestCaseBase.eval(String.format(GET, toString(), fun.substring(0,1).toUpperCase() + fun.substring(1)));
	}
	
	public WebElement toWebElement() {
		WebElement element = TestCaseBase.driver().findElement(ZSSBy.javascript("return " + toString()));
		if (element == null) {
			throw new RuntimeException("can't find element.");
		}
		return element;
	}
	
	public Object fun(String name) {
		return TestCaseBase.eval(String.format(RETURN_FUN, toString(), name));
	}
	
	public String getResult() {
		final String script = toString();
		js.clear();
		return script;
	}
	
	public Object getProperty(String property) {
		return TestCaseBase.eval("return " + toString() + "." + property);
	}
	
	public String toString() {
		StringBuilder sb = new StringBuilder();
		sb.append(selector + ".");
		for (String string : js) {
			sb.append(string).append('.');
		}
		if (sb.length() > 0 && sb.charAt(sb.length() - 1) == '.') {
			sb.deleteCharAt(sb.length() - 1);
		}
		return sb.toString();
	}
	
	public String getText() {
		return toWebElement().getText();
	}
}
