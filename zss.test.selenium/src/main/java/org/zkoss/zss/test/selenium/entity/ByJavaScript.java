/* ByJavaScript.java

	Purpose:
		
	Description:
		
	History:
		Mon, Jul 21, 2014  5:55:41 PM, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

 */
package org.zkoss.zss.test.selenium.entity;

import java.io.Serializable;
import java.util.List;
import java.util.Iterator;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.SearchContext;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.internal.WrapsDriver;

import com.google.common.base.Preconditions;
import com.google.common.base.Predicates;
import com.google.common.collect.Iterables;
import com.google.common.collect.Lists;

/**
 * @author RaymondChao
 */
public class ByJavaScript extends By implements Serializable {

	private static final long serialVersionUID = 1L;
	private final String script;

	public ByJavaScript(String script) {
		Preconditions.checkNotNull(script, "Cannot find elements with a null JavaScript expression.");
		this.script = script;
	}

	public List<WebElement> findElements(SearchContext context) {
		JavascriptExecutor excutor = getJavascriptExecutor(context);
		Object response = excutor.executeScript(script);
		List<WebElement> elements = getElementList(response);

		if (context instanceof WebElement) {
			filterByAncestor(elements, (WebElement) context);
		}
		return elements;
	}
	
	public WebElement findElement(SearchContext context) {
		List<WebElement> elements = findElements(context);
		if (elements.isEmpty()) {
			throw new NullPointerException("Cannot find any element");
		}
		return elements.get(0);
	}

	private static JavascriptExecutor getJavascriptExecutor(SearchContext context) {
		if (context instanceof JavascriptExecutor) {
			return (JavascriptExecutor) context;
		}
		if (context instanceof WrapsDriver) {
			WebDriver driver = ((WrapsDriver) context).getWrappedDriver();
			Preconditions.checkState(driver instanceof JavascriptExecutor, "This WebDriver doesn't support JavaScript.");
			return (JavascriptExecutor) driver;
		}
		throw new IllegalStateException("We can't invoke JavaScript from this context.");
	}

	@SuppressWarnings("unchecked")
	private static List<WebElement> getElementList(Object response) {
		if (response == null) {
			return Lists.newArrayList();
		}
		if (response instanceof WebElement) {
			return Lists.newArrayList((WebElement) response);
		}
		if (response instanceof List) {
			Preconditions.checkArgument(
					Iterables.all((List<?>) response, Predicates.instanceOf(WebElement.class)),
					"The JavaScript query returned something that isn't a WebElement.");
			return (List<WebElement>) response; // cast is checked as far as we can tell
		}
		throw new IllegalArgumentException("The JavaScript query returned something that isn't a WebElement.");
	}

	private static void filterByAncestor(List<WebElement> elements, WebElement ancestor) {
		Iterator<WebElement> iterator = elements.iterator();
		while (iterator.hasNext()) {
			WebElement elem = iterator.next();
			while (!elem.equals(ancestor) && !"html".equals(elem.getTagName())) {
				elem = elem.findElement(By.xpath("./.."));
			}
			if (!elem.equals(ancestor)) {
				iterator.remove();
			}
		}
	}

	public String toString() {
		return "By.javaScript: " + script;
	}
}