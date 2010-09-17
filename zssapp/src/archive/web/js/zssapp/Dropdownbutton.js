
/**
 * A button.
 * <p>Events: onOpen, onClick
 */
zssapp.Dropdownbutton = zk.$extends(zul.LabelImageWidget, {
	bind_: function () {
		this.$supers('bind_', arguments);
		this._cacheDpWidth();
		this._setClickable();
		zWatch.listen({onFloatUp: this});
	},
	/**
	 * Check whether server has listen onClick event or not, if not, set no clickable CSS   
	 */
	_setClickable: function () {
		if (!this.isListen('onClick', {asapOnly: true}))
			jq(this.$n()).addClass(this.getZclass() + '-noclk');
	},
	/**
	 * Calculate the drop down area width and cache it. 
	 */
	_cacheDpWidth: function () {
		var cnt = this.$n('cnt'),
			paddings = margins = 0;
		for (var n = this.$n().firstChild, i = 0; n ; n = n.firstChild) {
			paddings += zk.parseInt(jq(n).css('padding-right'));
			margins += zk.parseInt(jq(n).css('margin-right'));
			if (n == cnt)
				break;
		}
		this._dpWidth =  paddings + margins + zk.parseInt(jq(this.$n('btn')).css('width'));
	},
	unbind_: function () {
		this._dpWidth = null;
		zWatch.unlisten({onFloatUp: this});
		this.$supers('unbind_', arguments);
	},
	onFloatUp: function (ctl) {
		var wgt = ctl.origin;
		if (!zUtl.isAncestor(this.parent, ctl.origin))
			this._rmActive();
	},
	/**
	 * Base on mouse position, return click area
	 * <p> Return false (0) if click at main area
	 * <p> Return true (1) if click at dropdown arrow area
	 * @param mouse evt
	 * @return int
	 */
	_mouseAtDrop: function (evt) {
		var n = this.$n(),
			zkn = zk(n),
			w = zkn.offsetWidth(),
			x = evt.domEvent.clientX - zkn.revisedOffset()[0];
		return (x > (w - this._dpWidth));
	},
	onChildAdded_: function (child) {
		this.$supers('onChildAdded_', arguments);
		if (child.$instanceof(zul.wgt.Popup) || child.$instanceof(zul.menu.Menupopup))
			this.popup = child;
	},
	onChildRemoved_: function (child) {
		this.$supers('onChildRemoved_', arguments);
		if (child == this.popup)
			this.popup = null;
	},
	doClick_: function (evt) {
		if (!this._mouseAtDrop(evt)) {
			this.fire('onClick');
			this.$super('doClick_', evt, true);
		} else {
			var p = this.popup
			if (p && !p.isOpen()) {
				p.open(this, 0, 'after_start');
			}
			this.fire('onDropdown');
		}
	},
	doMouseOver_: function (evt) {
		var zcls = this.getZclass(),
			$n = jq(this.$n());
		
		if (this.isListen('onClick', {asapOnly: true})) {
			$n.addClass(zcls  + '-over');
			if (this._mouseAtDrop(evt))
				$n.addClass(zcls + '-dpover');
		} else
			$n.addClass(zcls + '-dpover');
	},
	_rmActive: function () {
		var zcls = this.getZclass();
		jq(this.$n()).removeClass(zcls + '-over').removeClass(zcls + '-dpover');
	},
	doMouseOut_: function (evt) {
		var p = this.popup;
		if (p && p.isOpen())
			return;
		this._rmActive();
	},
    getZclass: function() {
		var zcs = this._zclass;
		return zcs != null ? zcs: "z-dpbutton";
    }
});