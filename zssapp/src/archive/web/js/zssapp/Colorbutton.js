
zssapp.Colorbutton = zk.$extends(zkex.inp.Colorbox, {
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
		this.$n().style.backgroundColor = '#E3F9FF';
	},
	doMouseOut_: function () {
		this.$n().style.backgroundColor = '';
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
			colorNode = this.$n('currcolor');
		if (colorNode) {
			var color = this._currColor;
			colorNode.style.backgroundColor = color.getHex();
			//jq(n).css('backgroundColor', '');
		}
	},
	//override
	onShow: _cb,
    getZclass: function() {
		var zcs = this._zclass;
		return zcs != null ? zcs: "z-colorbtn";
    }
});