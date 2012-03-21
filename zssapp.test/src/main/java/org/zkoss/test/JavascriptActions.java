/* JavascriptActions.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 4:34:30 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import java.util.Collections;
import java.util.Set;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;

import com.google.common.collect.ImmutableSet;
import com.google.inject.Inject;

/**
 * 
 * @author sam
 *
 */
public class JavascriptActions extends Actions {

	final protected WebDriver webDriver;
	
	final protected JavascriptExecutor javascriptExecutor;
	
	@Inject
	public JavascriptActions (WebDriver webDriver) {
		super(webDriver);
		this.webDriver = webDriver;
		javascriptExecutor = (JavascriptExecutor)webDriver;
	}
	
	public JavascriptActions click(JQuery target) {
		mouseClick(target, MouseButton.LEFT);
		return this;
	}
	
	public JavascriptActions ctrlCopy(JQuery target) {
		int keyCode = Keycode.C.intValue();
		ctrlKeyDown(target, keyCode);
		ctrlKeyUp(target, keyCode);
		return this;
	}
	
	public JavascriptActions ctrlCut(JQuery target) {
		int keyCode = Keycode.X.intValue();
		ctrlKeyDown(target, keyCode);
		ctrlKeyUp(target, keyCode);
		return this;
	}
	
	public JavascriptActions ctrlPaste(JQuery target) {
		int keyCode = Keycode.V.intValue();
		ctrlKeyDown(target, keyCode);
		ctrlKeyUp(target, keyCode);
		return this;
	}
	
	public JavascriptActions ctrlClear(JQuery target) {
		int keyCode = Keycode.D.intValue();
		ctrlKeyDown(target, keyCode);
		ctrlKeyUp(target, keyCode);
		return this;
	}
	
	public JavascriptActions ctrlFontBold(JQuery target) {
		int keyCode = Keycode.B.intValue();
		ctrlKeyDown(target, keyCode);
		ctrlKeyUp(target, keyCode);
		return this;
	}
	
	public JavascriptActions ctrlFontItalic(JQuery target) {
		int keyCode = Keycode.I.intValue();
		ctrlKeyDown(target, keyCode);
		ctrlKeyUp(target, keyCode);
		return this;
	}
	
	public JavascriptActions ctrlFontUnderline(JQuery target) {
		int keyCode = Keycode.U.intValue();
		ctrlKeyDown(target, keyCode);
		ctrlKeyUp(target, keyCode);
		return this;
	}
	
	public JavascriptActions enter(JQuery target) {
		int keyCode = Keycode.ENTER.intValue();
		keyDown(target, keyCode);
		keyUp(target, keyCode);
		return this;
	}
	
	public JavascriptActions esc(JQuery target) {
		int keyCode = Keycode.ESC.intValue();
		keyDown(target, keyCode);
		keyUp(target, keyCode);
		return this;
	}
	
	public JavascriptActions delete(JQuery target) {
		int keyCode = Keycode.DELETE.intValue();
		keyDown(target, keyCode);
		keyUp(target, keyCode);
		return this;
	}

	public JavascriptActions ctrlKeyDown(JQuery target, int combinationKeyCode) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.KEYDOWN, combinationKeyCode, MouseButton.IGNORE, new CompundKey[]{CompundKey.CTRL}, 0, 0)));
		return this;
	}
	
	public JavascriptActions ctrlKeyUp(JQuery target, int combinationKeyCode) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.KEYUP, combinationKeyCode, MouseButton.IGNORE, new CompundKey[]{CompundKey.CTRL}, 0, 0)));
		return this;
	}
	
	public JavascriptActions shiftKeyUp(JQuery target, int combinationKeyCode) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.KEYUP, combinationKeyCode, MouseButton.IGNORE, new CompundKey[]{CompundKey.SHIFT}, 0, 0)));
		return this;
	}
	
	public JavascriptActions shiftKeyDown(JQuery target, int combinationKeyCode) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.KEYDOWN, combinationKeyCode, MouseButton.IGNORE, new CompundKey[]{CompundKey.SHIFT}, 0, 0)));
		return this;
	}
	
	public JavascriptActions mouseDown(JQuery target, MouseButton type, CompundKey...compundKeys) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.MOUSEDOWN, -1, type, compundKeys, 0, 0)));
		return this;
	}
	
	public JavascriptActions mouseUp(JQuery target, MouseButton type, CompundKey...compundKeys) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.MOUSEUP, -1, type, compundKeys, 0, 0)));
		return this;
	}
	
	public JavascriptActions mouseClick(JQuery target, MouseButton type, CompundKey...compundKeys) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.CLICK, -1, type, compundKeys, 0, 0)));
		return this;
	}
	
	public JavascriptActions contextClick(JQuery target) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.CONTEXTMENU, -1, MouseButton.RIGHT, null, 0, 0)));
		return this;
	}
	
	public JavascriptActions mouseOver(JQuery target, MouseButton type) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.MOUSEOVER, -1, type, null, 0, 0)));
		return this;
	}
	
	public JavascriptActions mouseMove(JQuery target, MouseButton type) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.MOUSEMOVE, -1, type, null, 0, 0)));
		return this;
	}
	
	public JavascriptActions mouseMove(JQuery target, MouseButton type, int xOffset, int yOffset) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.MOUSEMOVE, -1, type, null, xOffset, yOffset)));
		return this;	
	}
	
	public JavascriptActions mouseMove(int xOffset, int yOffset, JQuery target, MouseButton type) {
		return this;
	}
	
	public JavascriptActions keyDown(JQuery target, int keyCode) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.KEYDOWN, keyCode, MouseButton.IGNORE, null, 0, 0)));
		return this;
	}
	
	public JavascriptActions keyUp(JQuery target, int keyCode) {
		action.addAction(new ScriptAction(triggerEventScript(target, EventType.KEYUP, keyCode, MouseButton.IGNORE, null, 0, 0)));
		return this;
	}

	private String triggerEventScript(JQuery target, EventType eventType, int combinationKeyCode, MouseButton type, CompundKey[] compundKeys, int xOffset, int yOffset) {
		final String eventName = eventType.toString();
		Set<CompundKey> keys = compundKeys == null ? Collections.EMPTY_SET : ImmutableSet.copyOf(compundKeys);
		
		String script = "var evt = jq.Event('%eventName');" + 
			" evt.keyCode = %evtKeyCode; evt.metaKey = %evtMetaKey; evt.ctrlKey = %evtCtrlKey; evt.shiftKey = %evtShiftKey;" + 
			" evt.which = %evtWhich;" + "jq(%domTarget).trigger(evt);";
		
		script = script.replace("%domTarget", target.elementScript());
		script = script.replace("%eventName", eventName);
		if (combinationKeyCode < 0)
			script = script.replace("evt.keyCode = %evtKeyCode;", "");
		else
			script = script.replace("%evtKeyCode", "" + combinationKeyCode);
		int which = type.getWhich();
		if (which < 0) {
			script = script.replace("evt.which = %evtWhich;", "");
		} else {
			String offsetScript = "var ofs = " + target.zk().revisedOffsetScript();
			script = script.replace("%evtWhich;", "" + which + "; " + offsetScript + " evt.pageX=ofs[0]+" + xOffset 
					+";evt.pageY=ofs[1]+" + yOffset + ";evt.clientX=ofs[0]+ " + xOffset + ";evt.clientY=ofs[1]+ " + yOffset + ";");
		}
		boolean ctrlKey = keys.contains(CompundKey.CTRL),
				shiftKey = keys.contains(CompundKey.SHIFT);
		script = script.replace("%evtCtrlKey", Boolean.valueOf(ctrlKey).toString());
		script = script.replace("%evtShiftKey", Boolean.valueOf(shiftKey).toString());
		script = script.replace("%evtMetaKey", Boolean.valueOf(ctrlKey || shiftKey).toString());
		return script;
	}
	
	private class ScriptAction implements Action {
		String script;
		
		public ScriptAction(String script) {
			this.script = script;
		}
		
		public void perform() {
			javascriptExecutor.executeScript(script);
		}
	}
}
