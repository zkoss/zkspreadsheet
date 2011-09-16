function (out) {
	if (!this.getSheetId()) { //no sheet at all
		out.push('<div', this.domAttrs_(), '></div>');
		return;
	}
	var uuid = this.uuid,
		rowBegin = this._rowBegin,
		rowEnd = this._rowEnd,
		rowOuter = this._rowOuter,
		colBegin = this._colBegin,
		colEnd = this._colEnd,
		cellOuter = this._cellOuter,
		cellInner = this._cellInner,
		celltext = this._celltext,
		hidecolhead = this.isColumnHeadHidden(),
		hiderowhead = this.isRowHeadHidden(),
		topHeaderOuter = this._topHeaderOuter,
		topHeaderInner = this._topHeaderInner,
		columntitle = this._columntitle,
		leftHeaderOuter = this._leftHeaderOuter,
		leftHeaderInner = this._leftHeaderInner,
		rowtitle = this._rowtitle,
		rowfreezestr = this.getRowFreeze(),
		rowfreeze = rowfreezestr == null ? -1 : zk.parseInt(rowfreezestr),
		columnfreezestr = this.getColumnFreeze(),
		columnfreeze = columnfreezestr == null ? -1 : zk.parseInt(columnfreezestr),
		dataPanel = this._dataPanel,
		topHeaderHiddens = this._topHeaderHiddens,
		leftHeaderHiddens = this._leftHeaderHiddens;

	out.push('<div', this.domAttrs_(), '><textarea id="', uuid, '-fo" class="zsfocus"></textarea>',
			'<div id="', uuid, '-mask" class="zssmask" zs.t="SMask"><div class="zssmask2"><div id="', uuid, '-masktxt" class="zssmasktxt" align="center"></div></div></div>', 
			'<div id="', uuid, '-sp" class="zsscroll" zs.t="SScrollpanel">',
			'<div id="', uuid, '-dp" class="zsdata" zs.t="SDatapanel" ', dataPanel,' z.skipdsc="true">',
			'<div id="', uuid, '-datapad" class="zsdatapad"></div>',
			'<div id="', uuid, '-block" zs.t="SBlock" class="zsblock">');
	for (var r = rowBegin; r <= rowEnd; r++) {
		out.push('<div ', rowOuter[r],' zs.t="SRow">');
		
		for (var c = colBegin; c <= colEnd; c++) {
			out.push('<div zs.t="SCell" ', cellOuter[r][c], ' ><div ', cellInner[r][c], '>', zk.ie6_ || zk.ie7_ ? '<div style="left:0px;position:absolute;width:100%;">' + celltext[r][c] + '</div>' : celltext[r][c], '</div></div>');
		}
		out.push('</div>');
	}
	out.push('</div>',
			'<div id="', uuid, '-select" class="zsselect" zs.t="SSelect"><div id="', uuid, '-selecti" class="zsselecti" zs.t="SSelInner"></div><div class="zsseldot" zs.t="SSelDot"></div></div>',
			'<div id="', uuid, '-selchg" class="zsselchg" zs.t="SSelChg"><div id="', uuid, '-selchgi" class="zsselchgi"></div></div>',
			'<div id="', uuid, '-focmark" class="zsfocmark" zs.t="SFocus"><div id="', uuid, '-focmarki" class="zsfocmarki"></div></div>',
			'<div id="', uuid, '-highlight" class="zshighlight" zs.t="SHighlight"><div id="', uuid, '-highlighti" ,class="zshighlighti" zs.t="SHlInner"></div></div>',
			'</div>',
			'<textarea id="', uuid, '-eb" class="zsedit" zs.t="SEditbox"></textarea>',
			'<div id="', uuid, '-wp" class="zswidgetpanel" zs.t="SWidgetpanel"></div></div>',
			'<div id="', uuid, '-top" class="zstop zsfztop" zs.t="STopPanel" z.skipdsc="true">',
			'<div id="', uuid, '-topi" class="zstopi" >',
			'<div id="', uuid, '-tophead" class="zstophead" z.hide="', hidecolhead, '">');

	if (!hidecolhead) {
		for (var i = colBegin; i <= colEnd; i++) {
			out.push('<div zs.t="STheader" ', topHeaderOuter[i], '><div ', topHeaderInner[i], '>', columntitle[i],
					'</div></div><div class="zshboun"><div class="zshbouni" zs.t="SBoun"></div>');
			if (topHeaderHiddens[i])
				out.push('<div class="zshbounw" zs.t="SBoun2"></div>');
			out.push('</div>');
		}
	}

	out.push('</div>');
	if (rowfreeze >= 0) {
		out.push('<div id="', uuid, '-topblock" class="zstopblock">');
		for (var r = 0; r <= rowfreeze; r++) {
			out.push('<div ', rowOuter[r], ' zs.t="SRow">');
			for (var c = 0; c <= colEnd; c++) {
				out.push('<div zs.t="SCell" ', cellOuter[r][c], '><div ', cellInner[r][c], '>', celltext[r][c], '</div></div>');
			}
			out.push('</div>');
		}
		out.push('</div>',
				'<div class="zsselect" zs.t="SSelect"><div class="zsselecti" zs.t="SSelInner"></div><div class="zsseldot" zs.t="SSelDot"></div></div>',
				'<div class="zsselchg" zs.t="SSelChg"><div class="zsselchgi"></div></div>',
				'<div class="zsfocmark" zs.t="SFocus"><div class="zsfocmarki"></div></div>',
				'<div class="zshighlight" zs.t="SHighlight"><div class="zshighlighti"></div></div>');
	}
	out.push('</div>',
			'</div>',
			'<div id="', uuid, '-left" class="zsleft zsfzleft" zs.t="SLeftPanel" z.skipdsc="true">',
			'<div id="', uuid,'-leftpad" class="zsleftpad"></div>',
			'<div id="', uuid,'-lefti" class="zslefti">',
			'<div id="', uuid,'-lefthead" class="zslefthead" z.hide="', hiderowhead, '">');

	if (!hiderowhead) {
		for (var r = rowBegin; r <= rowEnd; r++) {
			out.push('<div zs.t="SLheader" ', leftHeaderOuter[r], '><div ', leftHeaderInner[r], '>', rowtitle[r], 
					'</div></div><div class="zsvboun"><div class="zsvbouni" zs.t="SBoun" ></div>');
			if (leftHeaderHiddens[r])
				out.push('<div class="zsvbounw" zs.t="SBoun2" ></div>');
			out.push('</div>');
		}
	}
	out.push('</div>');
	if (columnfreeze >= 0) {
		out.push('<div id="', uuid, '-leftblock" class="zsleftblock">');
		for (var r = 0; r <= rowEnd; r++) {
			out.push('<div ', rowOuter[r], 'zs.t="SRow">');
			for (var c = 0; c <= columnfreeze; c++) {
				out.push('<div zs.t="SCell" ', cellOuter[r][c], '><div ', cellInner[r][c], '>', celltext[r][c], '</div></div>');
			}
			out.push('</div>');
		}
		out.push('</div>',
				'<div class="zsselect" zs.t="SSelect"><div class="zsselecti" zs.t="SSelInner"></div><div class="zsseldot" zs.t="SSelDot"></div></div>',
				'<div class="zsselchg" zs.t="SSelChg"><div class="zsselchgi"></div></div>',
				'<div class="zsfocmark" zs.t="SFocus"><div class="zsfocmarki"></div></div>',
				'<div class="zshighlight" zs.t="SHighlight"><div class="zshighlighti"></div></div>');
	}
	out.push('</div></div><span id="', uuid, '-sinfo" class="zsscrollinfo"><span class="zsscrollinfoinner"></span></span>',
			'<span id="', uuid, '-info" class="zsinfo"><span class="zsinfoinner"></span></span>',
			'<div id="', uuid, '-co" class="zscorner zsfzcorner" zs.t="SCorner" z.skipdsc="true">');
	if (columnfreeze >= 0) {
		out.push('<div class="zscornertop">',
				'<div class="zscornertopi">',
				'<div class="zstophead" z.hide="', hidecolhead,'">');
		if (!hidecolhead) {
			for (var c = 0; c <= columnfreeze; c++) {
				out.push('<div zs.t="STheader" ', topHeaderOuter[c], '><div ', topHeaderInner[c], '>', columntitle[c], '</div></div><div class="zshboun"><div class="zshbouni" zs.t="SBoun" ></div></div>');
			}
		}
		out.push('</div></div></div>');
	}
	if (rowfreeze >= 0) {
		out.push('<div class="zscornerleft">',
				'<div class="zscornerpad"></div>',
				'<div class="zscornerlefti">',
				'<div class="zslefthead" z.hide="', hiderowhead, '">');
		if (!hiderowhead) {
			for (var r = 0; r <= rowfreeze; r++)
				out.push('<div zs.t="SLheader" ', leftHeaderOuter[r], '><div ', leftHeaderInner[r], '>', rowtitle[r], '</div></div><div class="zsvboun"><div class="zsvbouni" zs.t="SBoun" ></div></div>');
		}
		out.push('</div></div></div>');
	}
	
	if (rowfreeze >= 0 && columnfreeze >= 0) {
		out.push('<div class="zscornerblock">');
		for (var r = 0; r <= rowfreeze; r++) {
			out.push('<div ', rowOuter[r], 'zs.t="SRow">');
			 for (var c = 0; c <= columnfreeze; c++) {
				 out.push('<div zs.t="SCell" ', cellOuter[r][c], '><div ', cellInner[r][c], '>', celltext[r][c], '</div></div>');
			 }
			 out.push('</div>');
		}
		out.push('</div>',
			'<div class="zsselect" zs.t="SSelect"><div class="zsselecti" zs.t="SSelInner"></div><div class="zsseldot" zs.t="SSelDot"></div></div>',
			'<div class="zsselchg" zs.t="SSelChg"><div class="zsselchgi"></div></div>',
			'<div class="zsfocmark" zs.t="SFocus"><div class="zsfocmarki"></div></div>',
			'<div class="zshighlight" zs.t="SHighlight"><div class="zshighlighti"></div></div>');
	}
	out.push('<div class="zscorneri" ></div></div>');
    for(var w = this.firstChild; w ; w = w.nextSibling )
        w.redraw(out);
    out.push('</div>');
}