/* Colorbutton.js

{{IS_NOTE
	Purpose:

	Description:

	History:
		Sep 21, 2010 6:27:01 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHTRC
}}IS_RIGHT
*/
(function () {

zssapp.Simplecolorbutton = zk.$extends(zul.Widget, {
	_open: false,
	_colors: ["#000000","#993300","#333300","#003300","#000080",
	          "#003366","#660066","#333333","#800000","#FF8080",
	          "#808000","#008000","#008080","#0000FF","#800080",
	          "#969696","#FF0000","#FF6600","#FFFF66","#99CC00",
	          "#33CCCC","#0066CC","#6666FD","#808080","#FF00FF",
	          "#FF9900","#FFFF00","#00FF00","#00FFFF","#00CCFF",
	          "#CC99FF","#C0C0C0","#FF99CC","#FFCB90","#FFFF99",
	          "#CCFFCC","#CCFFFF","#99CCFF","#9999FF","#FFFFFF"],
	$define: {
		/** Sets the image URI.
		 * @param String image the URI of the image
		 */
		/** Returns the image URI.
		 * <p>Default: null.
		 * @return String
		 */
		image: function (v) {
			var n = this.getImageNode();
			if (n) n.src = v || '';
		},
		/**
		 * Sets the color
		 * @param color in #RRGGBB format (hexdecimal).
		 */
		/**
		 * Returns the color
		 * @return string
		 */
		color: function (hex) {
			var c = this.$n('currcolor');
			if (c)
				c.style.backgroundColor = hex;
		}
	},
	getImageNode: function () {
		return this.$n('btn');
	},
	bind_: function () {
		this.$supers('bind_', arguments);
		this.domListen_(this.$n(), 'onClick', '_doBtnClick');
		zWatch.listen({onFloatUp: this});

		//restore attribute to DOM elements
		this.setImage(this._image, {force: true});
		this.setColor(this._color, {force: true});
	},
	unbind_: function () {
		this.domUnlisten_(this.$n(), 'onClick', '_doBtnClick');
		zWatch.unlisten({onFloatUp: this});

		this.$supers(PalettePop, 'unbind_', arguments);
	},
	onFloatUp: function (ctl) {
		if (!zUtl.isAncestor(this.parent, ctl.origin))
			this.closePopup();
	},
	doMouseOver_: _onCell = function (evt) {
		var target = evt.domTarget,
			zcls = this.getZclass();
		if (jq(target).parents('.' + zcls + '-pp')[0]) {
			jq(target).parents('.' + zcls + '-cell')
			[evt.name == 'onMouseOver' ? 'addClass' : 'removeClass'](zcls + '-cell-over');
		} else
			jq(this.$n())[evt.name == 'onMouseOver' ? 'addClass' : 'removeClass'](this.getZclass() + '-over');
	},
	doMouseOut_: _onCell,
	doClick_: function (evt) {
		var t = evt.domTarget,
			$t = jq(evt.domTarget),
			zcls = this.getZclass();
		if ($t.attr('class').indexOf(zcls + '-cell') >= 0) {
			var hex = $t.children('i').text();
			this.fire("onChange", {color: hex}, {toServer: true});

			this.setColor(hex);
			this.closePopup();
		}
	},
	_doBtnClick: function (evt) {
		this._open ? this.closePopup() : this.openPopup();
	},
	openPopup: function () {
		this._open = true;
		var node = this.$n(),
			pp = this.$n("pp");

		pp.style.position = "absolute";
		pp.style.overflow = "auto";
		pp.style.display = "block";
		pp.style.zIndex = "88000";

		jq(pp).zk.makeVParent();
		zk(pp).position(node, this._getPosition());
	},
	_getPosition: function () {
		var parent = this.parent;
		if (!parent) return;
		if (parent.$instanceof(zul.wgt.Toolbar))
			return 'vertical' == parent.getOrient() ? 'end_before' : 'after_start';
		return 'after_start';
	},
	closePopup: function () {
		this._open = false;
		var node = this.$n(),
			pp = this.$n("pp");
		jq(pp).zk.undoVParent();
	},
	/**
	 * Returns the image DOM element
	 */
	getImageNode: function () {
		return this.$n('btn');
	},
    getZclass: function() {
		var zcs = this._zclass;
		return zcs != null ? zcs: "z-colorbtn";
    }
});
})();