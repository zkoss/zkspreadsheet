package org.zkoss.zss.essential;

import java.io.Serializable;
import java.util.Map;

import org.dom4j.IllegalAddException;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.Selectors;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.NRange;
import org.zkoss.zss.api.NRanges;
import org.zkoss.zss.api.model.NFont.Boldweight;
import org.zkoss.zss.api.model.NFont.Underline;
import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.api.ui.NSpreadsheet;
import org.zkoss.zss.essential.ToolbarCtrl.ClipInfo.Type;
import org.zkoss.zss.essential.util.ClientUtil;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.impl.XulElement;

public class ToolbarCtrl extends SelectorComposer<Component> {
	private static final long serialVersionUID = 1L;

	private static final String ON_INIT_SPREADSHEET = "onInitSpreadsheet";

	String ssid;
	String sspath;
	NSpreadsheet nss;

	@Wire
	Combobox fontNameBox;
	ListModelList<FontName> fontNameList;
	
	@Wire
	Combobox fontSizeBox;
	ListModelList<Integer> fontSizeList;

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		Map arg = Executions.getCurrent().getArg();
		ssid = (String) arg.get("spreadsheetId");
		sspath = (String) arg.get("spreadsheetPath");

		comp.addEventListener(ON_INIT_SPREADSHEET, new EventListener<Event>() {
			public void onEvent(Event event) throws Exception {
				getSelf().removeEventListener(ON_INIT_SPREADSHEET, this);
				postInitSpreadsheet();
			}
		});
		Events.postEvent(new Event(ON_INIT_SPREADSHEET, comp));

		String script = "focusSS(this,'" + ssid + "')";
		for (Component c : Selectors.find(comp, "toolbarbutton")) {
			((XulElement) c).setWidgetListener("onClick", script);
		}
		for (Component c : Selectors.find(comp, "combobox")) {
			((XulElement) c).setWidgetListener("onClick", script);
			((XulElement) c).setWidgetListener("onSelect", script);
		}
	}

	public void postInitSpreadsheet() {
		Component comp = getSelf();

		Spreadsheet ss = (Spreadsheet) comp.getFellowIfAny(ssid);
		if (ss == null && sspath != null) {
			IdSpace o = comp.getSpaceOwner();
			if (o != null) {
				ss = (Spreadsheet) Path.getComponent(o, sspath);
			}
		}
		if (ss == null && sspath != null) {
			ss = (Spreadsheet) Path.getComponent(sspath);
		}

		if (ss == null) {
			throw new IllegalAddException(
					"spreadsheet component not found with id " + ssid);
		}

		nss = new NSpreadsheet(ss);
		ssid = null;

		// initial font family
		fontNameList = new ListModelList<FontName>();
		fontNameList.add(new FontName("Arial", "Arial",
				"arial"));
		fontNameList.add(new FontName("Courier New", "Courier New",
				"courier-new"));
		fontNameList.add(new FontName("Times New Roman", "Times New Roman",
				"times-new-roman"));
		fontNameList.add(new FontName("MS Sans Serif", "MS Sans Serif",
				"ms-sans-serif"));

		fontNameBox.setModel(fontNameList);
		
		
		fontSizeList = new ListModelList<Integer>();
		fontSizeList.add(12);
		fontSizeList.add(14);
		fontSizeList.add(18);
		fontSizeList.add(24);
		fontSizeList.add(36);
		fontSizeList.add(48);

		fontSizeBox.setModel(fontSizeList);
	}

	ClipInfo clipinfo;

	static class ClipInfo {
		enum Type {
			COPY, CUT
		}

		Type type;
		NSheet sheet;
		Rect rect;

		public ClipInfo(NSheet sheet, Rect rect, Type type) {
			this.type = type;
			this.sheet = sheet;// TODO should I keep sheet instance?consider it
								// again.
			this.rect = rect;
		}
	}

	private void clearClipboard() {
		clipinfo = null;
		nss.setHighlight(null);
	}

	private void setClipboard(NSheet sheet, Rect rect, ClipInfo.Type type) {
		clipinfo = new ClipInfo(sheet, rect, type);
		nss.setHighlight(rect);
	}

	@Listen("onClick=#paste")
	public void doPaste() {
		if (clipinfo == null)
			return;

		Rect rect = nss.getSelection();
		if (rect == null)
			return;

		// check if in the same book only
		if (nss.getBook().getSheeteetIndex(clipinfo.sheet) < 0) {
			clearClipboard();
			return;
		}

		NRange src = NRanges.range(clipinfo.sheet, clipinfo.rect.getTop(),
				clipinfo.rect.getLeft(), clipinfo.rect.getBottom(),
				clipinfo.rect.getRight());

		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());

		boolean r;
		if (clipinfo.type == Type.CUT) {
			r = CellOperationUtil.cut(src,dest);
		}else{
			r = CellOperationUtil.copy(src,dest);
		}
		
		if (!r) {
			if (dest.isAnyCellProtected()) {
				// show protected message.
				ClientUtil.showWarn("Cann't paste to a protected sheet/area");
			} else {
				// TODO another reason?
				ClientUtil.showWarn("Cann't paste, reason unknow");
			}
			return;
		}
		if (clipinfo.type == Type.CUT) {
			clearClipboard();
		}

	}

	@Listen("onClick=#copy")
	public void doCopy() {
		Rect rect = nss.getSelection();
		setClipboard(nss.getSelectedSheet(), rect, ClipInfo.Type.COPY);
	}

	@Listen("onClick=#cut")
	public void doCut() {
		Rect rect = nss.getSelection();
		setClipboard(nss.getSelectedSheet(), rect, ClipInfo.Type.CUT);
	}

	static public class FontName implements Serializable {
		private static final long serialVersionUID = 1L;
		String displayStyle;
		String displayName;
		String fontName;

		public FontName(String fontName, String displayName,
				String displayStyle) {
			this.fontName = fontName;
			this.displayName = displayName;
			this.displayStyle = displayStyle;
		}

		public String getDisplayStyle() {
			return displayStyle;
		}
		public String getDisplayName() {
			return displayName;
		}

		public String getFontName() {
			return fontName;
		}
	}

	@Listen("onSelect=#fontNameBox")
	public void doFontName() {
		FontName font = fontNameList.getSelection().iterator().next();
		if (font == null) {
			return;
		}
		String fontName = font.getFontName();
		//
		Rect rect = nss.getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		//protection issue?
		CellOperationUtil.applyFontName(dest, fontName);
	}
	
	@Listen("onSelect=#fontSizeBox")
	public void doFontSize() {
		Integer size = fontSizeList.getSelection().iterator().next();
		if (size == null) {
			return;
		}
		
		//
		Rect rect = nss.getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		//protection issue?
		CellOperationUtil.applyFontSize(dest,size.shortValue());
	}
	
	@Listen("onClick=#fontBold")
	public void doFontBold() {
		Rect rect = nss.getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		NRange first = NRanges.range(nss.getSelectedSheet(), rect.getTop(),rect.getLeft());
		
		//toggle and apply bold of first cell to dest
		Boldweight bw = first.getGetter().getCellStyle().getFont().getBoldweight();
		if(Boldweight.BOLD.equals(bw)){
			bw = Boldweight.NORMAL;
		}else{
			bw = Boldweight.BOLD;
		}
		
		CellOperationUtil.applyFontBoldweight(dest, bw);	
	}

	
	@Listen("onClick=#fontItalic")
	public void doFontItalic() {
		Rect rect = nss.getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		NRange first = NRanges.range(nss.getSelectedSheet(), rect.getTop(),rect.getLeft());
		
		//toggle and apply bold of first cell to dest
		boolean italic = !first.getGetter().getCellStyle().getFont().isItalic();
		CellOperationUtil.applyFontItalic(dest, italic);	
	}
	
	@Listen("onClick=#fontStrike")
	public void doFontStrike() {
		Rect rect = nss.getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		NRange first = NRanges.range(nss.getSelectedSheet(), rect.getTop(),rect.getLeft());
		
		//toggle and apply bold of first cell to dest
		boolean strikeout = !first.getGetter().getCellStyle().getFont().isStrikeout();
		CellOperationUtil.applyFontStrikeout(dest, strikeout);
	}
	
	
	@Listen("onClick=#fontUnderline")
	public void doFontUnderline() {
		Rect rect = nss.getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		NRange first = NRanges.range(nss.getSelectedSheet(), rect.getTop(),rect.getLeft());
		
		//toggle and apply bold of first cell to dest
		Underline underline = first.getGetter().getCellStyle().getFont().getUnderline();
		if(Underline.NONE.equals(underline)){
			underline = Underline.SINGLE;
		}else{
			underline = Underline.NONE;
		}
		
		CellOperationUtil.applyFontUnderline(dest, underline);	
	}

	// <space sclass="sep-v" />
	// <!-- font -->
	// <combobox id="fontfamilyBox" style="width:105px"/>
	// <combobox id="fontsizeBox" style="width:45px"/>
	// <toolbarbutton id="fontBold" image="~./zss/img/edit-bold.png"/>
	// <toolbarbutton id="fontItalic" image="~./zss/img/edit-italic.png"/>
	// <toolbarbutton id="fontUnderline" image="~./zss/img/edit-underline.png"/>
	// <toolbarbutton id="fontStrike" image="~./zss/img/edit-strike.png"/>
}
