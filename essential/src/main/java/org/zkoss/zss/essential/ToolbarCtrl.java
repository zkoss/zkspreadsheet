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
import org.zkoss.zss.api.NRange.ApplyBorderType;
import org.zkoss.zss.api.NRange.ApplyBorderLineStyle;
import org.zkoss.zss.api.NRanges;
import org.zkoss.zss.api.model.NCellStyle;
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
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menupopup;
import org.zkoss.zul.Toolbarbutton;
import org.zkoss.zul.impl.XulElement;

public class ToolbarCtrl extends SelectorComposer<Component> {
	private static final long serialVersionUID = 1L;

	private static final String ON_INIT_SPREADSHEET = "onInitSpreadsheet";

	String ssid;
	String sspath;
	NSpreadsheet nss;

	
	@Wire
	Toolbarbutton paste;
	@Wire
	Toolbarbutton pasteMenu;
	@Wire
	Menupopup pastePopup;
	
	@Wire
	Combobox fontNameBox;
	ListModelList<FontName> fontNameList;
	
	@Wire
	Combobox fontSizeBox;
	ListModelList<Integer> fontSizeList;
	
//	@Wire
//	Colorbox fontColorbox;
//	@Wire
//	Colorbox fillColorbox;
	
	@Wire
	Toolbarbutton fontColorMenu;
	@Wire
	Menupopup fontColorPopup;
	@Wire
	Toolbarbutton fillColorMenu;
	@Wire
	Menupopup fillColorPopup;
	@Wire
	Menu fontColor;
	@Wire
	Menu fillColor;	
	
	
	
	@Wire
	Toolbarbutton alignMenu;
	@Wire
	Menupopup alignPopup;
	
	
	
	@Wire
	Toolbarbutton borderMenu;
	@Wire
	Menupopup borderPopup;
	@Wire
	Menu borderColor;	
	
	
	@Wire
	Toolbarbutton mergeMenu;
	@Wire
	Menupopup mergePopup;

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
		
		clearClipboard();
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
		pasteMenu.setDisabled(true);
		paste.setDisabled(true);
	}

	private void setClipboard(NSheet sheet, Rect rect, ClipInfo.Type type) {
		clipinfo = new ClipInfo(sheet, rect, type);
		nss.setHighlight(rect);
	}
	
	enum PasteType{
		ALL,
		VALUE,
		FORMULA,
		ALL_NO_BORDER,
		TRANSPORT
	}

	@Listen("onClick=#pasteMenu")
	public void doPasteMenu(){
		pastePopup.open(pasteMenu);
	}

	@Listen("onClick=#paste")
	public void doPaste() {
		doPaste0(PasteType.ALL);
	}
	@Listen("onClick=#pasteValue")
	public void doPasteValue() {
		doPaste0(PasteType.VALUE);
	}
	@Listen("onClick=#pasteFormula")
	public void doPasteFormula() {
		doPaste0(PasteType.FORMULA);
	}
	@Listen("onClick=#pasteAllNOBorder")
	public void doAllNOBorder() {
		doPaste0(PasteType.ALL_NO_BORDER);
	}
	@Listen("onClick=#pasteTranspose")
	public void doPasteTransport() {
		doPaste0(PasteType.TRANSPORT);
	}
	
	
	private void doPaste0(PasteType type) {
		if (clipinfo == null)
			return;

		Rect rect = getSelection();
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

		boolean r = true;
		if (clipinfo.type == Type.CUT) {
			r = CellOperationUtil.cut(src,dest);
		}else{
			switch(type){
			case ALL:
				r = CellOperationUtil.paste(src,dest);
				break;
			case VALUE:
				r = CellOperationUtil.pasteValue(src, dest);
				break;
			case FORMULA:
				r = CellOperationUtil.pasteFormula(src, dest);
				break;
			case ALL_NO_BORDER:
				r = CellOperationUtil.pasteAllExceptBorder(src, dest);
				break;
			case TRANSPORT:
				r = CellOperationUtil.pasteTranspose(src, dest);
				break;
				
			}
		}
		
		if (!r) {
			if (dest.isProtected()) {
				// show protected message.
				ClientUtil.showWarn("Cann't paste to a protected sheet/area");
			} else {
				if (clipinfo.type == Type.CUT && src.isProtected()) {
					ClientUtil.showWarn("Cann't cut from a protected sheet/area");
				}
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
		Rect rect = getSelection();
		setClipboard(nss.getSelectedSheet(), rect, ClipInfo.Type.COPY);
		pasteMenu.setDisabled(false);
		paste.setDisabled(false);
	}

	@Listen("onClick=#cut")
	public void doCut() {
		Rect rect = getSelection();
		setClipboard(nss.getSelectedSheet(), rect, ClipInfo.Type.CUT);
		//TODO should disable some past-special toolbar button
		pasteMenu.setDisabled(true);
		paste.setDisabled(false);
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
		Rect rect = getSelection();
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
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		//protection issue?
		CellOperationUtil.applyFontSize(dest,size.shortValue());
	}
	
	@Listen("onClick=#fontBold")
	public void doFontBold() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		NRange first = dest.getFirst();
		
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
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		NRange first = dest.getFirst();
		
		//toggle and apply bold of first cell to dest
		boolean italic = !first.getGetter().getCellStyle().getFont().isItalic();
		CellOperationUtil.applyFontItalic(dest, italic);	
	}
	
	@Listen("onClick=#fontStrike")
	public void doFontStrike() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		NRange first = dest.getFirst();
		
		//toggle and apply bold of first cell to dest
		boolean strikeout = !first.getGetter().getCellStyle().getFont().isStrikeout();
		CellOperationUtil.applyFontStrikeout(dest, strikeout);
	}
	
	
	@Listen("onClick=#fontUnderline")
	public void doFontUnderline() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		NRange first = dest.getFirst();
		
		//toggle and apply bold of first cell to dest
		Underline underline = first.getGetter().getCellStyle().getFont().getUnderline();
		if(Underline.NONE.equals(underline)){
			underline = Underline.SINGLE;
		}else{
			underline = Underline.NONE;
		}
		
		CellOperationUtil.applyFontUnderline(dest, underline);	
	}
	
	
	private Rect getSelection(){
		Rect rect = nss.getSelection();
		//TODO re-format rect before zss support well rows,colums selection handling
		int r = rect.getRight();
		int b = rect.getBottom();
		rect = new Rect(rect.getLeft(), rect.getTop(),
				(r <= nss.getMaxcolumns()) ? r : nss.getMaxcolumns(),
				(b <= nss.getMaxrows()) ? b : nss.getMaxrows());

		return rect;
	}
	
	//Color
	@Listen("onClick=#fontColorMenu")
	public void doFontColorMenu(){
		fontColorPopup.open(fontColorMenu);
	}
	@Listen("onClick=#fillColorMenu")
	public void doFillColorMenu(){
		fillColorPopup.open(fillColorMenu);
	}
	
	@Listen("onChange=#fontColor")
	public void doFontColor(){
		String htmlColor = fontColor.getContent(); //'#HEX-RGB'
		if(htmlColor==null){
			return;
		}
		htmlColor = htmlColor.substring(htmlColor.lastIndexOf("#"));

		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		CellOperationUtil.applyFontColor(dest, htmlColor);		
	}
	
	@Listen("onChange=#fillColor")
	public void doFillColor(){
		String htmlColor = fillColor.getContent(); //'#HEX-RGB'
		if(htmlColor==null){
			return;
		}
		htmlColor = htmlColor.substring(htmlColor.lastIndexOf("#"));

		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		CellOperationUtil.applyCellColor(dest, htmlColor);		
	}
	
	//Align
	@Listen("onClick=#alignMenu")
	public void doAlignMenu(){
		alignPopup.open(alignMenu);
	}

	@Listen("onClick=#alignTop")
	public void doAlignTop() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		CellOperationUtil.applyCellVerticalAlignment(dest,NCellStyle.VerticalAlignment.TOP);
	}
	@Listen("onClick=#alignMiddle")
	public void doAlignMiddle() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		CellOperationUtil.applyCellVerticalAlignment(dest,NCellStyle.VerticalAlignment.CENTER);
	}
	@Listen("onClick=#alignBottom")
	public void doAlignBottom() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		CellOperationUtil.applyCellVerticalAlignment(dest,NCellStyle.VerticalAlignment.BOTTOM);
	}
	
	
	@Listen("onClick=#alignLeft")
	public void doAlignLeft() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		CellOperationUtil.applyCellAlignment(dest,NCellStyle.Alignment.LEFT);
	}
	@Listen("onClick=#alignCenter")
	public void doAlignCenter() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		CellOperationUtil.applyCellAlignment(dest,NCellStyle.Alignment.CENTER);
	}
	@Listen("onClick=#alignRight")
	public void doAlignRight() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		CellOperationUtil.applyCellAlignment(dest,NCellStyle.Alignment.RIGHT);
	}
	
	
	//Border
	@Listen("onClick=#borderMenu")
	public void doBorderMenu(){
		borderPopup.open(borderMenu);
	}
	@Listen("onClick=#borderAll")
	public void onBorderAll(){
		doBorder0(ApplyBorderType.FULL,ApplyBorderLineStyle.MEDIUM);
	}
	@Listen("onClick=#borderNo")
	public void onBorderNo(){
		doBorder0(ApplyBorderType.FULL,ApplyBorderLineStyle.NONE);
	}
	
	@Listen("onClick=#borderBottom")
	public void onBorderBottom(){
		doBorder0(ApplyBorderType.EDGE_BOTTOM,ApplyBorderLineStyle.MEDIUM);
	}
	
	@Listen("onClick=#borderTop")
	public void onBorderTop(){
		doBorder0(ApplyBorderType.EDGE_TOP,ApplyBorderLineStyle.MEDIUM);
	}
	
	@Listen("onClick=#borderLeft")
	public void onBorderLeft(){
		doBorder0(ApplyBorderType.EDGE_LEFT,ApplyBorderLineStyle.MEDIUM);
	}
	
	@Listen("onClick=#borderRight")
	public void onBorderRight(){
		doBorder0(ApplyBorderType.EDGE_RIGHT,ApplyBorderLineStyle.MEDIUM);
	}
	
	@Listen("onClick=#borderOutside")
	public void onBorderOutside(){
		doBorder0(ApplyBorderType.OUTLINE,ApplyBorderLineStyle.MEDIUM);
	}
	
	@Listen("onClick=#borderInside")
	public void onBorderInside(){
		doBorder0(ApplyBorderType.INSIDE,ApplyBorderLineStyle.MEDIUM);
	}
	
	@Listen("onClick=#borderInsideHorizontal")
	public void onBorderInsideHor(){
		doBorder0(ApplyBorderType.INSIDE_HORIZONTAL,ApplyBorderLineStyle.MEDIUM);
	}
	
	@Listen("onClick=#borderInsideVertical")
	public void onBorderInsideVer(){
		doBorder0(ApplyBorderType.INSIDE_VERTICAL,ApplyBorderLineStyle.MEDIUM);
	}
	
	
	
	private void doBorder0(ApplyBorderType type,ApplyBorderLineStyle style){
		String htmlColor = borderColor.getContent(); //'#HEX-RGB'
		if(htmlColor==null){
			return;
		}
		htmlColor = htmlColor.substring(htmlColor.lastIndexOf("#"));
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		CellOperationUtil.applyBorder(dest,type, style, htmlColor);
	}
	
	// Wrap
	@Listen("onClick=#wrapTextMenu")
	public void doWrapTExt() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		
		NRange first = dest.getFirst();
		
		//toggle and apply 
		boolean wrapped = first.getGetter().getCellStyle().isWrapText();
		
		wrapped = !wrapped;
		
		CellOperationUtil.applyCellWrapText(dest, wrapped);	
	}
	
	// Merge
	@Listen("onClick=#mergeMenu")
	public void doMergeMenu() {
		mergePopup.open(mergeMenu);
	}

	@Listen("onClick=#mergeCenter")
	public void onMergeCenter() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		CellOperationUtil.toggleMergeCenter(dest);
	}

	@Listen("onClick=#mergeAcross")
	public void onMergeAcross() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		CellOperationUtil.merge(dest, true);
	}

	@Listen("onClick=#mergeAll")
	public void onMergeAll() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		CellOperationUtil.merge(dest, false);
	}

	@Listen("onClick=#unMerge")
	public void onUnMerge() {
		Rect rect = getSelection();
		NRange dest = NRanges.range(nss.getSelectedSheet(), rect.getTop(),
				rect.getLeft(), rect.getBottom(), rect.getRight());
		CellOperationUtil.unMerge(dest);
	}
}
