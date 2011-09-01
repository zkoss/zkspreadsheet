zk.afterLoad('zss', function () {
	var APP = 'app',
		FILE = 'file',
		CP = 'copyPaste',
		ID = 'insertDelete',
		FILTER = 'filter',
		SORT = 'sort',
		FONT = 'font',
		BORDER = 'border',
		ALIGN = 'align',
		CELL = 'cell',
		RC = 'rowcolumn',
		SHEET = 'sheet',
		F = 'formula',
		N = 'numberFormat',
		VIEW = 'view',
		HELP = 'help',
		EVENT = 'event',
		STYLE = 'style';
	
	function _getMonitorInfo(id) {
		var info = zssAct.monitorIds[id],
			newInfo = null;
		if (info) {
			var now = jq.now(),
				interval = now - zssAct.time;
			zssAct.time = now;
			newInfo = {action: info.action, label: info.label};
			newInfo.value = interval; 
		}
		return newInfo;
	}
	
	function _appendLabel(info, extra) {
		info.label = info.label + '=' + extra + ';';
		//return src += ('=' + extra + ';');
	}
	window.zssAct = {
			//last action timestampe
			category: 'zssdemo',
			time: jq.now(),
			monitorIds: {
				/*app*/
				'app0': {action: APP, label: 'ZK Spreadsheet App'},
				'app2': {action: APP, label: 'ZK Spreadsheet H1N1 Demo'},
				/*ss event*/
				'onStopEditing': {action: EVENT, label: 'onStopEditing'},
				/*file*/
				'newFile': {action: FILE, label: 'newEmptySpreadsheet'}, 
				'deleteFile': {action: FILE, label: 'deleteSpreadsheet'},
				'exportPdf': {action: FILE, label: 'exportPDF'},
				'exportToPDFBtn': {action: FILE, label: 'exportPDF'},
				'exportHtml': {action: FILE, label: 'exportHtml'},
				'exportExcel': {action: FILE, label: 'exportExcel'},
				'closeBtn': {action: FILE, label: 'closeSpreadsheet'},
				/*copy and paste*/
				'cut': {action: CP, label: 'cut'},
				'cutBtn': {action: CP, label: 'cut'},
				'copyBtn': {action: CP, label: 'copy'},
				'copy': {action: CP, label: 'copy'},
				'pasteDropdownBtn': {action: CP, label: 'paste'},
				'paste': {action: CP, label: 'paste'},
				'pasteFormula': {action: CP, label: 'pasteFormula'},
				'pasteValue': {action: CP, label: 'pasteValue'},
				'pasteAllExcpetBorder': {action: CP, label: 'pasteAllExcpetBorder'},
				'pasteTranspose': {action: CP, label: 'pasteTranspose'},
				'pasteSpecial': {action: CP, label: 'pasteSpecialDialog'},
				/*sort*/
				'sortAscending': {action: SORT, label: 'sortAscending'},
				'sortDescending': {action: SORT, label: 'sortDescending'},
				'customSort': {action: SORT, label: 'customSort'},
				/*filter*/
				'filter': {action: FILTER, label: 'applyFilter'},
				'clearFilter': {action: FILTER, label: 'clearFilter'},
				'reapplyFilter': {action: FILTER, label: 'reapplyFilter'},
				/*insertDelete*/
				'shiftCellRight': {action: ID, label: 'shiftCellRight'},
				'shiftCellDown': {action: ID, label: 'shiftCellDown'},
				'insertRow': {action: ID, label: 'insertRows'},
				'insertEntireRow': {action: ID, label: 'insertRows'},
				'insertColumn': {action: ID, label: 'insertColumns'},
				'insertEntireColumn': {action: ID, label: 'insertColumns'},
				'shiftCellLeft': {action: ID, label: 'shiftCellLeft'},
				'shiftCellUp': {action: ID, label: 'shiftCellUp'},
				'deleteRow': {action: ID, label: 'deleteRows'},
				'deleteEntireRow': {action: ID, label: 'deleteRows'},
				'deleteColumn': {action: ID, label: 'deleteColumns'},
				'deleteEntireColumn': {action: ID, label: 'deleteColumns'},
				'insertImage': {action: ID, label: 'insertImage'},
				'insertSheet': {action: ID, label: 'insertSheet'},
				/*style*/
				'fontFamily': {action: FONT, label: 'setFontFamily'},
				'fontSize': {action: FONT, label: 'setFontSize'},
				'fontColorBtn': {action: FONT, label: 'setFontColor'},
				'boldBtn': {action: FONT, label: 'setFontBold'},
				'italicBtn': {action: FONT, label: 'setFontItalic'},
				'underlineBtn': {action: FONT, label: 'setFontUnderline'},
				'strikethroughBtn': {action: FONT, label: 'setFontStrikethrough'},
				/*border*/
				'borderBtn': {action: BORDER, label: 'setBorder=Bottom'},
				'bottomBorder': {action: BORDER, label: 'setBorder=Bottom'},
				'topBorder': {action: BORDER, label: 'setBorder=Top'},
				'leftBorder': {action: BORDER, label: 'setBorder=Left'},
				'rightBorder': {action: BORDER, label: 'setBorder=Right'},
				'noBorder': {action: BORDER, label: 'setBorder=No'},
				'fullBorder': {action: BORDER, label: 'setBorder=All'},
				'outsideBorder': {action: BORDER, label: 'setBorder=Outside'},
				'insideBorder': {action: BORDER, label: 'setBorder=Inside'},
				'insideHorizontalBorder': {action: BORDER, label: 'setBorder=InsideHorizontal'},
				'insideVerticalBorder': {action: BORDER, label: 'setBorder=InsideVertical'},
				/*align*/
				'valignTop': {action: ALIGN, label: 'valign=Top'},
				'valignMiddle': {action: ALIGN, label: 'valign=Middle'},
				'valignBottom': {action: ALIGN, label: 'valign=Bottom'},
				'halignLeft': {action: ALIGN, label: 'halign=Left'},
				'halignCenter': {action: ALIGN, label: 'halign=Center'},
				'halignRight': {action: ALIGN, label: 'halign=Right'},
				/*view*/
				'viewFormulaBar': {action: VIEW, label: 'viewFormulaBar'},
				'unfreezeRows': {action: VIEW, label: 'unfreezeRows'},
				'freezeRow1': {action: VIEW, label: 'freezeRow=1'},
				'freezeRow2': {action: VIEW, label: 'freezeRow=2'},
				'freezeRow3': {action: VIEW, label: 'freezeRow=3'},
				'freezeRow4': {action: VIEW, label: 'freezeRow=4'},
				'freezeRow5': {action: VIEW, label: 'freezeRow=5'},
				'freezeRow6': {action: VIEW, label: 'freezeRow=6'},
				'freezeRow7': {action: VIEW, label: 'freezeRow=7'},
				'freezeRow8': {action: VIEW, label: 'freezeRow=8'},
				'freezeRow9': {action:VIEW, label: 'freezeRow=9'},
				'freezeRow10': {action: VIEW, label: 'freezeRow=10'},
				'unfreezeCols': {action: VIEW, label: 'unfreezeCols'},
				'freezeCol1': {action: VIEW, label: 'freezeCol=1'},
				'freezeCol2': {action: VIEW, label: 'freezeCol=2'},
				'freezeCol3': {action: VIEW, label: 'freezeCol=3'},
				'freezeCol4': {action: VIEW, label: 'freezeCol=4'},
				'freezeCol5': {action: VIEW, label: 'freezeCol=5'},
				'freezeCol6': {action: VIEW, label: 'freezeCol=6'},
				'freezeCol7': {action: VIEW, label: 'freezeCol=7'},
				'freezeCol8': {action: VIEW, label: 'freezeCol=8'},
				'freezeCol9': {action: VIEW, label: 'freezeCol=9'},
				'freezeCol10': {action: VIEW, label: 'freezeCol=10'},
				/*help*/
				'openCheatsheet': {action: HELP, label: 'openCheatsheet'},
				'forum': {action: HELP, label: 'forum'},
				'book': {action: HELP, label: 'book'},
				/*cell*/
				'cellColorBtn': {action: CELL, label: 'setCellColor'},
				'wrapTextBtn': {action: CELL, label: 'wrapText'},
				'mergeCellBtn': {action: CELL, label: 'mergeCells'},
				'locked': {action: CELL, label: 'locked'},
				'insertHyperlinkBtn': {action: CELL, label: 'insertHyperlink'},
				'removeHyperlink': {action: CELL, label: 'removeHyperlink'},
				'clearContent': {action: CELL, label: 'clearContent'},
				'clearStyle': {action: CELL, label: 'clearStyle'},
				'clearAll': {action: CELL, label: 'clearContentAndStyle'},
				/*row column*/
				'columnWidth': {action: RC, label: 'setColumnWith'},
				'rowHeight': {action: RC, label: 'setRowHeight'},
				'hide': {action: RC, label: 'hide'},
				'unhide': {action: RC, label: 'unhide'},
				/*sheet*/
				'tabbox': {action: SHEET, label: 'selectSheet'},
				'shiftSheetLeft': {action: SHEET, label: 'shiftSheetLeft'},
				'shiftSheetRight': {action: SHEET, label: 'shiftSheetRight'},
				'deleteSheet': {action: SHEET, label: 'deleteSheet'},
				'renameSheet': {action: SHEET, label: 'renameSheet'},
				'protectSheet': {action: SHEET, label: 'protectSheet'},
				'gridlinesCheckbox': {action: SHEET, label: 'gridlines'},
				/*formula*/
				'formula': {action: F, label: 'insertFormulaDialog'},
				'insertFormula': {action: F, label: 'insertFormulaDialog'},
				'insertFormulaBtn': {action: F, label: 'insertFormulaDialog'},
				'searchTextbox': {action: F, label: 'searchFormula'},
				'searchBtn': {action: F, label: 'searchFormula'},
				'functionListbox': {action: F, label: 'selectFormula'},
				'categoryCombobox': {action: F, label: 'selectFormulaCategory'},
				'formulaEditor': {action: F, label: 'formulaEditor'},
				/*number format*/
				'format': {action: N, label: 'numberFormatDialog'},
				'numberFormat': {action: N, label: 'numberFormatDialog'},
				'mfn_category': {action: N, label: 'selectNumberFormatCategory'},
				'mfn_number': {action: N, label: 'formatNumber'},
				'mfn_currency': {action: N, label: 'formatCurrency'},
				'mfn_accounting': {action: N, label: 'formatAccounting'},
				'mfn_date': {action: N, label: 'formatDate'},
				'mfn_time': {action: N, label: 'formatTime'},
				'mfn_percentage': {action: N, label: 'formatPercentage'},
				'mfn_fraction': {action: N, label: 'formatFraction'},
				'mfn_scientific': {action: N, label: 'formatScientific'},
				'mfn_text': {action: N, label: 'formatText'},
				'mfn_special': {action: N, label: 'formatSpecial'}
			}
	}
	
	//tab
	zk.afterLoad('zul.tab', function () {
		var ORG_TABBOX_DOCLICK = zul.tab.Tab.prototype.doClick_;
		zul.tab.Tab.prototype.doClick_ = function (evt) {
			if (window.pageTracker) {
				var tid = this.getTabbox().id,
					info = _getMonitorInfo(tid);
				if (info && tid =='tabbox') {
					_appendLabel(info, this.getLabel());
					pageTracker._trackEvent(zssAct.category, info.action, info.label, info.value);
				}
			}
			return ORG_TABBOX_DOCLICK.apply(this, arguments);
		};
	});
	
	//combo, textbox
	zk.afterLoad('zul.inp', function () {
		var ORG_COMBOBOX_FIREX = zul.inp.Combobox.prototype.fireX;
		zul.inp.Combobox.prototype.fireX = function (evt) {
			if (window.pageTracker && evt.name == 'onSelect') {
				var id = this.id,
					info = _getMonitorInfo(id);
				if (info) {
					var label = info.label;
					if (label == 'setFontFamily' || label == 'setFontSize' || label == 'selectFormulaCategory') {
						 _appendLabel(info, this.getText());
						 pageTracker._trackEvent(zssAct.category, info.action, info.label, info.value);
					}
				}
			}
			return ORG_COMBOBOX_FIREX.apply(this, arguments);
		};
		
		var ORG_INPUT_FIREX = zul.inp.Textbox.prototype.fireX;
		zul.inp.Textbox.prototype.fireX = function (evt) {
			if (window.pageTracker && evt.name == 'onChange') {
				var id = this.id,
				info = _getMonitorInfo(id);
				if (info) {
					if (id == 'formulaEditor')
						_appendLabel(info, '\'' + this.getText() + '\'');
					else
						_appendLabel(info, this.getText());
					pageTracker._trackEvent(zssAct.category, info.action, info.label, info.value);
				}
			}
			return ORG_INPUT_FIREX.apply(this, arguments);
		}
	});
	
	//toolbar, checkbox,
	zk.afterLoad('zul.wgt', function (){
		
		var ORG_TOOLBAR_FIREX = zul.wgt.Toolbarbutton.prototype.fireX;
		zul.wgt.Toolbarbutton.prototype.fireX = function (evt) {
			if (window.pageTracker && evt.name == 'onClick') {
				var id = this.id,
					info = _getMonitorInfo(id);
				if (info) {
					pageTracker._trackEvent(zssAct.category, info.action, info.label, info.value);
				}
			}
			return ORG_TOOLBAR_FIREX.apply(this, arguments);
		};
		
		var ORG_CHECKBOX_FIREX = zul.wgt.Checkbox.prototype.fireX;
		zul.wgt.Checkbox.prototype.fireX = function (evt) {
			if (window.pageTracker && evt.name == 'onCheck') {
				var id = this.id,
					info = _getMonitorInfo(id);
				if (info) {
					pageTracker._trackEvent(zssAct.category, info.action, info.label, info.value);
				}
			}
			return ORG_CHECKBOX_FIREX.apply(this, arguments);
		}
		//div
		var ORG_DIV_FIREX = zul.wgt.Div.prototype.fireX;
		zul.wgt.Div.prototype.fireX = function (evt) {
			if (window.pageTracker && evt.name == 'onClick') {
				var id = this.id,
					info = _getMonitorInfo(id);
				if (info) {
					pageTracker._trackEvent(zssAct.category, info.action, info.label, info.value);
				}
			}
			return ORG_DIV_FIREX.apply(this, arguments);
		}
	});
	
	//menuitem
	zk.afterLoad('zul.menu', function () {
		var ORG_MENUITEM_DOCLICK = zul.menu.Menuitem.prototype.doClick_;
		zul.menu.Menuitem.prototype.doClick_ = function (evt) {
			if (window.pageTracker) {
				var id = this.id,
					info = _getMonitorInfo(id);
				if (info) {
					pageTracker._trackEvent(zssAct.category, info.action, info.label, info.value);
				}
			}
			return ORG_MENUITEM_DOCLICK.apply(this, arguments);
		};
	});
	
	//dropdown button
	zk.afterLoad('zssapp', function () {
		var ORG_DROPDOWN_BUTTON_DOCLICK = zssapp.Dropdownbutton.prototype.fireX;
		zssapp.Dropdownbutton.prototype.fireX = function (evt) {
			if (window.pageTracker && evt.name == 'onClick') {
				var id = this.id,
					info = _getMonitorInfo(id);
				if (info) {
					pageTracker._trackEvent(zssAct.category, info.action, info.label, info.value);
				}
			}
			return ORG_DROPDOWN_BUTTON_DOCLICK.apply(this, arguments);
		}
	});
	//colorbox
	zk.afterLoad('zkex.inp', function (){
		var ORG_COLORBOX_FIREX = zkex.inp.Colorbox.prototype.fireX;
		zkex.inp.Colorbox.prototype.fireX = function (evt) {
			if (window.pageTracker && evt.name == 'onChange') {
				var id = this.id,
					info = _getMonitorInfo(id);
				if (info) {
					pageTracker._trackEvent(zssAct.category, info.action, info.label, info.value);
				}
			}
			return ORG_COLORBOX_FIREX.apply(this, arguments);
		};
	});
	
	//zul.sel.Listbox 
	zk.afterLoad('zul.sel', function () {
		var ORG_LB_FIREX = zul.sel.Listbox.prototype.fireX;
		zul.sel.Listbox.prototype.fireX = function (evt) {
			if (window.pageTracker && evt.name == 'onSelect') {
				var id = this.id,
					info = _getMonitorInfo(id);
				if (info) {
					if (this._selItems && this._selItems[0])
						_appendLabel(info, this._selItems[0].getLabel());
					pageTracker._trackEvent(zssAct.category, info.action, info.label, info.value);
				}
			}
			return ORG_LB_FIREX.apply(this, arguments);
		};
	});
});