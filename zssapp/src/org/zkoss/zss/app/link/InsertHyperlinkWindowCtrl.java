/* InsertHyperlinkWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 7, 2010 11:34:37 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.link;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashMap;

import org.zkoss.lang.Objects;
import org.zkoss.poi.ss.usermodel.Hyperlink;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.event.InputEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapps;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.SheetVisitor;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Button;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Include;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.SimpleTreeModel;
import org.zkoss.zul.SimpleTreeNode;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Tree;
import org.zkoss.zul.Treeitem;
import org.zkoss.zul.Window;
import org.zkoss.zul.api.Comboitem;

/**
 * @author Sam
 *
 */
public class InsertHyperlinkWindowCtrl extends GenericForwardComposer {

	private Dialog _insertHyperlinkDialog;
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
	
	private final static String WEBLINK_CONTENT_URI = Consts._Weblink_zul;
	private final static String DOCLINK_CONTENT_URI = Consts._Doclink_zul;
	private final static String MAILLINK_CONTENT_URI = Consts._Maillink_zul;
	private Include content;
	
	private Spreadsheet ss;
	private Rect selection;
	private boolean isCellHasDisplayString;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		ss = checkNotNull(Zssapps.getSpreadsheetFromArg(), "Spreadsheet is null");
		linkTypeBtns.put(webBtn, WEBLINK_CONTENT_URI);
		linkTypeBtns.put(docBtn, DOCLINK_CONTENT_URI);
		linkTypeBtns.put(mailBtn, MAILLINK_CONTENT_URI);
		
		setLinkType(webBtn);
	}
	
	public void onOpen$_insertHyperlinkDialog(ForwardEvent event) {
		selection = (Rect) event.getOrigin().getData();
		init();
		try {
			_insertHyperlinkDialog.setMode(Window.MODAL);
		} catch (InterruptedException e) {
		}
	}
	
	private void init() {
		String display = Utils.getRange(ss.getSelectedSheet(), selection.getTop(), selection.getLeft()).getEditText();
		isCellHasDisplayString = !"".equals(display);
		if (isCellHasDisplayString)
			displayHyperlink.setValue(display);

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
		
		Utils.setHyperlink(ss.getSelectedSheet(), selection.getTop(), selection.getLeft(), 
				getLinkTarget(), addr, getDisplay());

		_insertHyperlinkDialog.fireOnClose(null);
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
		mailAddr.addEventListener(Events.ON_OK, onOkEventListener);
		mailSubject.addEventListener(Events.ON_OK, onOkEventListener);
		mailAddr.focus();
	}
	
	private EventListener onOkEventListener = new EventListener(){
		@Override
		public void onEvent(Event evt) throws Exception {
			onClick$okBtn();
		}};
	
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
				Comboitem seld = addr.getSelectedItem();
				if (seld != null)
					displayHyperlink.setValue(seld.getLabel());
			}
		});
		addr.addEventListener(Events.ON_OK, onOkEventListener);
		addr.focus();
	}

	/**
	 * Returns the string of web page address, return null if component not found
	 * @return
	 */
	private String getWebAddress() {
		Combobox address = (Combobox)content.getFellow("addrCombobox");
		String val = address.getValue();
		return val == null ? "" : val;
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
		cellRef.focus();
	}
	private void buildDocumentTree(final Tree tree, final Textbox cellRef) {
		if (tree != null) {
			final Book book = ss.getBook();
			if (book == null) {
				return;
			}
			final ArrayList<SimpleTreeNode> nodes = new ArrayList<SimpleTreeNode>();
			Utils.visitSheets(book, new SheetVisitor(){
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