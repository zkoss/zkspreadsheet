<%--
ss.dsp

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		January 10, 2008 4:58:31 PM , Created by Dennis Chen
}}IS_NOTE

Copyright (C) 2008 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
--%><%@ taglib uri="http://www.zkoss.org/dsp/web/core" prefix="c" %>
<%@ taglib uri="http://www.zkoss.org/dsp/zk/core" prefix="z" %>
<%@ taglib uri="http://www.zkoss.org/zss/ui" prefix="s" %>

<c:set var="self" value="${requestScope.arg.self}"/>
<div class="zssheet" id="${self.uuid}"${self.outerAttrs}${self.innerAttrs} z.type="zss.ss.SSheet" zs.t="SSheet">
	<input type="textbox" id="${self.uuid}-fo" maxlength="0" class="zsfocus" /><div class="zssmask" zs.t="SMask"></div><div class="zssbusy"></div>
	<div id="${self.uuid}-sp" class="zsscroll" zs.t="SScrollpanel">
		<div id="${self.uuid}-dp" class="zsdata" zs.t="SDatapanel" ${self.extraCtrl.dataPanelAttrs} z.skipdsc="true">
			<div class="zsdatapad"></div>
			<div zs.t="SBlock" class="zsblock"><c:forEach begin="${s:getRowBegin(self)}" end="${s:getRowEnd(self)}" varStatus="rstatus">
				<div ${s:getRowOuterAttrs(self,rstatus.index)} zs.t="SRow"><%-- cell div must be side by side, no space between allowed --%>
				<c:forEach begin="${s:getColBegin(self)}" end="${s:getColEnd(self)}" varStatus="cstatus"><div zs.t="SCell" ${s:getCellOuterAttrs(self,rstatus.index,cstatus.index)}><div ${s:getCellInnerAttrs(self,rstatus.index,cstatus.index)}>${s:getCelltext(self,rstatus.index,cstatus.index)}</div></div></c:forEach>
				</div></c:forEach>
			</div>
			<div class="zsselect" zs.t="SSelect"><div class="zsselecti" zs.t="SSelInner"></div><div class="zsseldot" zs.t="SSelDot"></div></div>
			<div class="zsselchg" zs.t="SSelChg"><div class="zsselchgi"></div></div>
			<div class="zsfocmark" zs.t="SFocus"><div class="zsfocmarki"></div></div>
			<div class="zshighlight" zs.t="SHighlight"><div class="zshighlighti" zs.t="SHlInner"></div></div>
			<textarea id="${self.uuid}-eb" class="zsedit" zs.t="SEditbox"></textarea>
		</div>
		<div class="zswidgetpanel" zs.t="SWidgetpanel">
			${s:getWidgetContent(self)}
		</div>
	</div>
	<div id="${self.uuid}-top" class="zstop zsfztop" zs.t="STopPanel" z.skipdsc="true">
		<div id="${self.uuid}-topi" class="zstopi" >
			<div class="zstophead" z.hide="${self.hidecolumnhead}"><c:if test="${!self.hidecolumnhead}">
			<%-- cell div must be side by side, no space between allowed --%>
			<c:forEach begin="${s:getColBegin(self)}" end="${s:getColEnd(self)}" varStatus="status"><div zs.t="STheader" ${s:getTopHeaderOuterAttrs(self,status.index)}><div ${s:getTopHeaderInnerAttrs(self,status.index)}>${s:getColumntitle(self,status.index)}</div></div><div class="zshboun"><div class="zshbouni" zs.t="SBoun"></div></div></c:forEach>
			</c:if></div>
			<c:if test="${self.rowfreeze>=0}">
			<div class="zstopblock"><c:forEach begin="0" end="${self.rowfreeze}" varStatus="rstatus">
				<div ${s:getRowOuterAttrs(self,rstatus.index)} zs.t="SRow"><%-- cell div must be side by side, no space between allowed --%>
				<c:forEach begin="0" end="${s:getColEnd(self)}" varStatus="cstatus"><div zs.t="SCell" ${s:getCellOuterAttrs(self,rstatus.index,cstatus.index)}><div ${s:getCellInnerAttrs(self,rstatus.index,cstatus.index)}>${s:getCelltext(self,rstatus.index,cstatus.index)}</div></div></c:forEach>
				</div></c:forEach>
			</div>
			<div class="zsselect" zs.t="SSelect"><div class="zsselecti" zs.t="SSelInner"></div><div class="zsseldot" zs.t="SSelDot"></div></div>
			<div class="zsselchg" zs.t="SSelChg"><div class="zsselchgi"></div></div>
			<div class="zsfocmark" zs.t="SFocus"><div class="zsfocmarki"></div></div>
			<div class="zshighlight" zs.t="SHighlight"><div class="zshighlighti"></div></div>
			</c:if>
		</div>
	</div>
	<div id="${self.uuid}-left" class="zsleft zsfzleft" zs.t="SLeftPanel" z.skipdsc="true">
		<div class="zsleftpad"></div>
		<div id="${self.uuid}-lefti" class="zslefti">
			<div class="zslefthead" z.hide="${self.hiderowhead}"><c:if test="${!self.hiderowhead}">
			<c:forEach begin="${s:getRowBegin(self)}" end="${s:getRowEnd(self)}" varStatus="status">
			<div zs.t="SLheader" ${s:getLeftHeaderOuterAttrs(self,status.index)}><div ${s:getLeftHeaderInnerAttrs(self,status.index)}>${s:getRowtitle(self,status.index)}</div></div><div class="zsvboun"><div class="zsvbouni" zs.t="SBoun" ></div></div>
			</c:forEach>
			</c:if></div>
			<c:if test="${self.columnfreeze>=0}">
			<div class="zsleftblock"><c:forEach begin="0" end="${s:getRowEnd(self)}" varStatus="rstatus">
				<div ${s:getRowOuterAttrs(self,rstatus.index)} zs.t="SRow"><%-- cell div must be side by side, no space between allowed --%>
				<c:forEach begin="0" end="${self.columnfreeze}" varStatus="cstatus"><div zs.t="SCell" ${s:getCellOuterAttrs(self,rstatus.index,cstatus.index)}><div ${s:getCellInnerAttrs(self,rstatus.index,cstatus.index)}>${s:getCelltext(self,rstatus.index,cstatus.index)}</div></div></c:forEach>
				</div></c:forEach>
			</div>
			<div class="zsselect" zs.t="SSelect"><div class="zsselecti" zs.t="SSelInner"></div><div class="zsseldot" zs.t="SSelDot"></div></div>
			<div class="zsselchg" zs.t="SSelChg"><div class="zsselchgi"></div></div>
			<div class="zsfocmark" zs.t="SFocus"><div class="zsfocmarki"></div></div>
			<div class="zshighlight" zs.t="SHighlight"><div class="zshighlighti"></div></div>
			</c:if>
		</div>
	</div>
	<span class="zsscrollinfo"><span class="zsscrollinfoinner"></span></span>
	<span class="zsinfo"><span class="zsinfoinner"></span></span>
	<div id="${self.uuid}-co" class="zscorner zsfzcorner" zs.t="SCorner" z.skipdsc="true">
		<c:if test="${self.columnfreeze>=0}">
		<div class="zscornertop">
			<div class="zscornertopi">
				<div class="zstophead" z.hide="${self.hidecolumnhead}"><c:if test="${!self.hidecolumnhead}">
				<%-- cell div must be side by side, no space between allowed --%>
				<c:forEach begin="0" end="${self.columnfreeze}" varStatus="status"><div zs.t="STheader" ${s:getTopHeaderOuterAttrs(self,status.index)}><div ${s:getTopHeaderInnerAttrs(self,status.index)}>${s:getColumntitle(self,status.index)}</div></div><div class="zshboun"><div class="zshbouni" zs.t="SBoun" ></div></div></c:forEach>
				</c:if></div>
			</div>
		</div>
		</c:if>
		
		<c:if test="${self.rowfreeze>=0}">
		<div class="zscornerleft">
			<div class="zscornerpad"></div>
			<div class="zscornerlefti">
				<div class="zslefthead" z.hide="${self.hiderowhead}"><c:if test="${!self.hiderowhead}">
				<c:forEach begin="0" end="${self.rowfreeze}" varStatus="status">
				<div zs.t="SLheader" ${s:getLeftHeaderOuterAttrs(self,status.index)}><div ${s:getLeftHeaderInnerAttrs(self,status.index)}>${s:getRowtitle(self,status.index)}</div></div><div class="zsvboun"><div class="zsvbouni" zs.t="SBoun" ></div></div>
				</c:forEach>
				</c:if></div>
			</div>
		</div>
		</c:if>
		
		<c:if test="${self.rowfreeze>=0 && self.columnfreeze>=0}">
		<div class="zscornerblock"><c:forEach begin="0" end="${self.rowfreeze}" varStatus="rstatus">
			<div ${s:getRowOuterAttrs(self,rstatus.index)} zs.t="SRow"><%-- cell div must be side by side, no space between allowed --%>
			<c:forEach begin="0" end="${self.columnfreeze}" varStatus="cstatus"><div zs.t="SCell" ${s:getCellOuterAttrs(self,rstatus.index,cstatus.index)}><div ${s:getCellInnerAttrs(self,rstatus.index,cstatus.index)}>${s:getCelltext(self,rstatus.index,cstatus.index)}</div></div></c:forEach>
			</div></c:forEach>
		</div>
		<div class="zsselect" zs.t="SSelect"><div class="zsselecti" zs.t="SSelInner"></div><div class="zsseldot" zs.t="SSelDot"></div></div>
		<div class="zsselchg" zs.t="SSelChg"><div class="zsselchgi"></div></div>
		<div class="zsfocmark" zs.t="SFocus"><div class="zsfocmarki"></div></div>
		<div class="zshighlight" zs.t="SHighlight"><div class="zshighlighti"></div></div>
		</c:if>
		
		<div class="zscorneri" >
		</div>
	</div>
	<c:forEach var="child" items="${self.children}">
		${z:redraw(child, null)}
	</c:forEach>
</div>