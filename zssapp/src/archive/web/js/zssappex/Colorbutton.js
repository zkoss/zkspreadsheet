/* Colorbutton.js

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 21, 2010 6:27:01 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
(function () {

zssappex.Colorbutton = zk.$extends(zkex.inp.Colorbox, {
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
		}
	},
	doMouseOver_: function () {
		jq(this.$n()).addClass(this.getZclass() + '-over');
	},
	doMouseOut_: function () {
		jq(this.$n()).removeClass(this.getZclass() + '-over');
	},
	bind_: function () {
		this.$supers('bind_', arguments);
		var img = this.getImageNode();
		if (img)
			img.src = this.getImage();
	},
	//
	_doBtnClick: function (evt) {
		this.$supers('_doBtnClick', arguments);
		this.fire("onClick");
	},
	/**
	 * Returns the image DOM element
	 */
	getImageNode: function () {
		return this.$n('btn');
	},
	//override
	onSize: _cb = function () {
		var n = this.$n(),
			c = this.$n('currcolor');
		if (c)
			c.style.backgroundColor = this._currColor.getHex();
	},
	//override
	onShow: _cb,
    getZclass: function() {
		var zcs = this._zclass;
		return zcs != null ? zcs: "z-colorbtn";
    }
});
})();