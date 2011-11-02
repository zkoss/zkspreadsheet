function (out) {
	if (!this.getSheetId()) { //no sheet at all
		out.push('<div', this.domAttrs_(), '></div>');
		return;
	}
	var uuid = this.uuid,
		sheet = this.sheetCtrl,
		activeBlock = sheet.activeBlock,
		topPanel = sheet.tp,
		leftPanel = sheet.lp,
		cornerPanel = sheet.cp,
		hidecolhead = this.isColumnHeadHidden(),
		hiderowhead = this.isRowHeadHidden(),
		dataPanel = this._dataPanel;
	
	out.push('<div', this.domAttrs_(), '><textarea id="', uuid, '-fo" class="zsfocus"></textarea>',
			'<div id="', uuid, '-mask" class="zssmask" zs.t="SMask"><div class="zssmask2"><div id="', uuid, '-masktxt" class="zssmasktxt" align="center"></div></div></div>', 
			'<div id="', uuid, '-sp" class="zsscroll" zs.t="SScrollpanel">',
			'<div id="', uuid, '-dp" class="zsdata" zs.t="SDatapanel" ', dataPanel,' z.skipdsc="true">',
			'<div id="', uuid, '-datapad" class="zsdatapad"></div>');

	if (activeBlock)
		activeBlock.redraw(out);
	
	out.push(
			'<div id="', uuid, '-select" class="zsselect" zs.t="SSelect"><div id="', uuid, '-selecti" class="zsselecti" zs.t="SSelInner"></div><div class="zsseldot" zs.t="SSelDot"></div></div>',
			'<div id="', uuid, '-selchg" class="zsselchg" zs.t="SSelChg"><div id="', uuid, '-selchgi" class="zsselchgi"></div></div>',
			'<div id="', uuid, '-focmark" class="zsfocmark" zs.t="SFocus"><div id="', uuid, '-focmarki" class="zsfocmarki"></div></div>',
			'<div id="', uuid, '-highlight" class="zshighlight" zs.t="SHighlight"><div id="', uuid, '-highlighti" ,class="zshighlighti" zs.t="SHlInner"></div></div>',
			'</div>',
			'<textarea id="', uuid, '-eb" class="zsedit" zs.t="SEditbox"></textarea>',
			'<div id="', uuid, '-wp" class="zswidgetpanel" zs.t="SWidgetpanel"></div></div>');
	
	if (topPanel)
		topPanel.redraw(out);
	
	if (leftPanel)
		leftPanel.redraw(out);
	
	out.push('<span id="', uuid, '-sinfo" class="zsscrollinfo"><span class="zsscrollinfoinner"></span></span>',
			'<span id="', uuid, '-info" class="zsinfo"><span class="zsinfoinner"></span></span>');
	
	if (cornerPanel)
		cornerPanel.redraw(out);
	
    out.push('</div>');
}