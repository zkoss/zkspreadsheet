/* Button.js

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 12, 2012 7:07:27 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
(function () {

zss.ToolbarTabpanel = zk.$extends(zul.tab.Tabpanel, {
	$init: function (actions, wgt) {
		this.$supers(zss.ToolbarTabpanel, '$init', []);
		this._wgt = wgt;
		this._actions = actions;//array
	},
	createButtonsIfNotExist: function () {
		var tb = this.toolbar;
		if (!tb) {
			tb = this.toolbar = new zul.wgt.Toolbar({sclass: 'zstoolbar'});
			var btns = new zss.ButtonBuilder(this._wgt).addAll(this._actions).build();
			for (var i = 0, len = btns.length; i < len; i++) {
				var b = btns[i];
				if (b)
					tb.appendChild(b);
			}
			this.appendChild(tb);
			this.setDisabled(this._wgt.getActionDisabled());
		}
	},
	onShow: function () {
		this.createButtonsIfNotExist();
	},
	setDisabled: function (actions) {
		var tb = this.toolbar;
		if (tb && actions) {
			var btn = tb.firstChild;
			for (; btn; btn = btn.nextSibling) {
				if (!btn.setDisabled) {
					continue;
				}
				btn.setDisabled(actions);
			}	
		}
	}
})

zss.ToolbarTabbox = zk.$extends(zul.tab.Tabbox, {
	$o: zk.$void,
	$init: function (wgt) {
		this.$supers(zss.ToolbarTabbox, '$init', []);
		this._wgt = wgt;
		
		var labels = wgt._labelsCtrl,
			tbs = new zul.tab.Tabs(),
			homeTab = new zul.tab.Tab({
				label: labels.getHomePanel(),
				sclass: 'zstab-homePanel'
			}),
			insertTab = new zul.tab.Tab({
				label: labels.getInsertPanel(),
				onClick: this.proxy(this.onClickInsertTab),
				sclass: 'zstab-insertPanel'
			}),
			//TODO: formulaTab = new zul.tab.Tab({
			//	label: labels.getFormulaPanel(),
			//	onClick: this.proxy(this.onClickFormulaTab)
			//}),
			panels = new zul.tab.Tabpanels(),
			homePanel = this.homePanel = new zss.ToolbarTabpanel(zss.Buttons.HOME_DEFAULT, wgt),
			//TODO: formulaPanel = this.formulaPanel = new zul.tab.Tabpanel(zss.Buttons.FORMULA_DEFAULT),
			insertPanel = this.insertPanel = new zss.ToolbarTabpanel(zss.Buttons.INSERT_DEFAULT, wgt);
//			homeToolbar = new zul.wgt.Toolbar({height: '23px'});

		this.appendChild(tbs);
		this.appendChild(panels);
		
		tbs.appendChild(homeTab);
		tbs.appendChild(insertTab);
		//TODO: tbs.appendChild(formulaTab);

		homePanel.createButtonsIfNotExist();
		panels.appendChild(homePanel);
		panels.appendChild(insertPanel);
		//TODO: panels.appendChild(formulaPanel);
		
		this.setSelectedTab(homeTab);
		
		var tb = new zul.wgt.Toolbar();
		tb.appendChild(new zss.Toolbarbutton({
			$action: 'closeBook',
			tooltiptext: wgt._labelsCtrl.getCloseBook(),
			image: zk.ajaxURI('/web/zss/img/gray-cross.png', {au: true}),
			onClick: function () {
				wgt.fireToolbarAction('closeBook');
			}
		}));
		this.appendChild(tb);
	},
	onClickInsertTab: function () {
		this.insertPanel.createButtonsIfNotExist();
	},
	onClickFormulaTab: function () {
		this.formulaPanel.createButtonsIfNotExist();
	}
});

zss.Toolbar = zk.$extends(zul.layout.North, {
	$o: zk.$void,
	$init: function (wgt) {
		this.$supers(zss.Toolbar, '$init', []);
		this.setSize('53px');
		this.setBorder(0);
		this._wgt = wgt;
		
		this.appendChild(this.toolbarTabbox = new zss.ToolbarTabbox(wgt));
	},
	setDisabled: function (actions) {
		var tb = this.toolbarTabbox,
			panels = tb.getTabpanels(),
			panel = panels.firstChild;
		for (; panel; panel = panel.nextSibling) {
			panel.setDisabled(actions);
		}
	}
});
})();