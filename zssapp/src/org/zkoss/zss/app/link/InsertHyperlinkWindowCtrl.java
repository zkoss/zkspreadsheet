/* InsertHyperlinkWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 7, 2010 11:34:37 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.link;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashMap;

import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Sheet;
import org.zkoss.lang.Objects;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.InputEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.SheetVisitor;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Button;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Comboitem;
import org.zkoss.zul.Include;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.SimpleTreeModel;
import org.zkoss.zul.SimpleTreeNode;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Tree;
import org.zkoss.zul.Treeitem;

/**
 * @author Sam
 *
 */
public class InsertHyperlinkWindowCtrl extends GenericForwardComposer {
	
	private Button webBtn;
	private Button docBtn;
	private Button mailBtn;
	/**
	 * Current selected button
	 */
	private Button activeBtn;
	private LinkedHashMap<Button, String> linkTypeBtns = new LinkedHashMap<Button, String>(3);
	
	private Button okBtn;
	
	private Textbox displayHyperlink;
	
	private final static String WEBLINK_CONTENT_URI="/menus/hyperlink/webLink.zul";
	private final static String DOCLINK_CONTENT_URI="/menus/hyperlink/docLink.zul";
	private final static String MAILLINK_CONTENT_URI="/menus/hyperlink/mailLink.zul";
	private Include content;
	
	private Spreadsheet ss;
	
	private boolean isCellHasDisplayString;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		ss = (Spreadsheet)getParam("spreadsheet");
		if (ss == null)
			throw new UiException("Spreadsheet object is null");
		
		String display = Utils.getRange(ss.getSelectedSheet(), ss.getSelection().getTop(), ss.getSelection().getLeft()).getEditText();
		isCellHasDisplayString = !"".equals(display);
		if (isCellHasDisplayString)
			displayHyperlink.setValue(display);
		
		initPageContent();
	}
	
	private static Object getParam (String key) {
		return Executions.getCurrent().getArg().get(key);
	}
	
	/**
	 * Setup default page to web link
	 */
	private void initPageContent() {
		linkTypeBtns.put(webBtn, WEBLINK_CONTENT_URI);
		linkTypeBtns.put(docBtn, DOCLINK_CONTENT_URI);
		linkTypeBtns.put(mailBtn, MAILLINK_CONTENT_URI);
		
		setLinkType(webBtn);
	}

	public void onClick$webBtn() {
		setLinkType(webBtn);
	}
	public void onClick$docBtn() {
		setLinkType(docBtn);
	}
	public void onClick$mailBtn() {
		setLinkType(mailBtn);
	}
	
	public void onClick$okBtn() {
		String addr = getAddress();
		if ("".equals(addr)) {
			try {
				Messagebox.show("Please input address");
			} catch (InterruptedException e) {
			}
			return;
		}
		
		Utils.setHyperlink(ss.getSelectedSheet(), ss.getSelection().getTop(), ss.getSelection().getLeft(), 
				getLinkTarget(), addr, getDisplay());

		((Component)spaceOwner).detach();
	}
	
	private int getLinkTarget() {
		if (docBtn == activeBtn)
			return Hyperlink.LINK_DOCUMENT;
		else if (mailBtn == activeBtn)
			return Hyperlink.LINK_EMAIL;
		return Hyperlink.LINK_URL;
	}
	/**
	 * Returns link address
	 * @return
	 */
	private String getAddress() {
		if (docBtn == activeBtn)
			return getDocAddress();
		else if (mailBtn == activeBtn)
			return getMailAddress();
		return getWebAddress();
	}
	
	private String getDisplay() {
		String text = displayHyperlink.getValue();
		return "".equals(text) ? "" : text;
	} 
	
	private void setLinkType(Button btn) {
		activeBtn = btn;
		for (Iterator<Button> iter = linkTypeBtns.keySet().iterator(); iter.hasNext();) {
			Button b = iter.next();
			//set selected button disable
			b.setDisabled( Objects.equals(btn, b) );
		}
		content.setSrc( linkTypeBtns.get(btn) );
		
		if (!isCellHasDisplayString)
			displayHyperlink.setValue("");
		if (Objects.equals(docBtn, btn)) {
			initDoc();
		} else if (Objects.equals(webBtn, btn)) {
			initAddr();
		} else if (Objects.equals(mailBtn, btn)) {
			initMail();
		}
	}
	
	private void initMail() {
		if (isCellHasDisplayString)
			return;
		
		final Textbox mailAddr = (Textbox)content.getFellow("mailAddr");
		final Textbox mailSubject = (Textbox)content.getFellow("mailSubject");
		final String preAppend = "mailto:";
		mailAddr.addEventListener(Events.ON_CHANGING, new EventListener() {
			
			@Override
			public void onEvent(Event evt) throws Exception {
				String mail = mailSubject.getValue();
				String val = preAppend + ((InputEvent)evt).getValue() + 
							(mail != null && mail != "" ? "?subject=" + mailSubject.getValue() : "");
				displayHyperlink.setValue(val);
			}
		});
		
		mailSubject.addEventListener(Events.ON_CHANGING, new EventListener() {
			
			@Override
			public void onEvent(Event evt) throws Exception {
				String mail = mailAddr.getValue();
				if (mail != null  && mail != "")
					displayHyperlink.setValue(preAppend + mailAddr.getValue() + "?subject=" + ((InputEvent)evt).getValue());
			}
		});
	}

	private void initAddr() {
		if (isCellHasDisplayString)
			return;

		final Combobox addr = (Combobox)content.getFellow("addrCombobox");
		addr.addEventListener(Events.ON_CHANGING, new EventListener() {
			
			@Override
			public void onEvent(Event evt) {
				displayHyperlink.setValue(((InputEvent)evt).getValue());
			}
		});
		addr.addEventListener(Events.ON_SELECT, new EventListener() {
			
			@Override
			public void onEvent(Event evt) throws Exception {
				displayHyperlink.setValue(addr.getSelectedItem().getLabel());
			}
		});
	}

	/**
	 * Returns the string of web page address, return null if component not found
	 * @return
	 */
	private String getWebAddress() {
		Combobox address = (Combobox)content.getFellow("addrCombobox");
		Comboitem item = address.getSelectedItem();
		//TODO input in combobox
		return item != null ? item.getLabel() : "";
	}
	
	private String getDocAddress() {
		Tree sheetTree = (Tree)content.getFellow("refSheet");
		Treeitem sheet = sheetTree.getSelectedItem();
		if (sheet == null)
			return "";
		
		String sheetName = sheet.getLabel();
		String cell = ((Textbox)content.getFellow("cellRef")).getValue();
		return sheetName + "!" + cell;
	}
	
	private String getMailAddress() {
		String mailAddr = ((Textbox)content.getFellow("mailAddr")).getValue();
		if ("".equals(mailAddr))
			return "";
		
		String subject = ((Textbox)content.getFellow("mailSubject")).getValue();
		return "mailto:" + mailAddr + (("".equals(subject)) ? "" : "?subject=" + subject);
	}
	
	/**
	 * Returns the cell reference string, return null if component not found
	 * @return
	 */
	public String getCellRef() {
		Textbox cellRef = (Textbox)content.getFellowIfAny("cellRef");
		return cellRef != null ? cellRef.getText() : "";
	}
	

	private void initDoc() {
		final Tree tree = (Tree)content.getFellow("refSheet");
		final Textbox cellRef = (Textbox)content.getFellow("cellRef");
		
		buildDocumentTree(tree, cellRef);
		cellRef.addEventListener(Events.ON_CHANGING, new EventListener() {
			
			@Override
			public void onEvent(Event evt) throws Exception {
				displayHyperlink.setValue(tree.getSelectedItem().getLabel() + "!" + ((InputEvent)evt).getValue());
			}
		});
	}
	private void buildDocumentTree(final Tree tree, final Textbox cellRef) {
		if (tree != null) {
			final ArrayList<SimpleTreeNode> nodes = new ArrayList<SimpleTreeNode>();
			Utils.visitSheets(ss.getBook(), new SheetVisitor(){
				@Override
				public void handle(Sheet sheet) {
					nodes.add(new SimpleTreeNode(sheet.getSheetName(), Collections.EMPTY_LIST));
				}});
			
			/**
			 * TODO: use i-18n instead hardcode
			 */
			SimpleTreeNode root = new SimpleTreeNode("Cell Reference", nodes);
			SimpleTreeModel model = new SimpleTreeModel(root);
			tree.setModel(model);
			
			tree.addEventListener(Events.ON_SELECT, new EventListener() {
				
				@Override
				public void onEvent(Event evt) throws Exception {
					displayHyperlink.setValue(tree.getSelectedItem().getLabel() + "!" + cellRef.getValue());
				}
			});
		}
	}
	
	/**
	 * Returns the reference sheet name, return null if link type is not document
	 * @return
	 */
	public String getRefSheet() {
		Textbox refSheet = (Textbox)content.getFellowIfAny("refSheet");
		return refSheet != null ? refSheet.getText() : "";	
	}
}