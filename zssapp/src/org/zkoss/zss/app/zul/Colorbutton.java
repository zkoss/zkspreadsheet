/* Colorbutton.java

{{IS_NOTE
	Purpose:

	Description:

	History:
		Oct 19, 2010 2:50:50 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.zul;

import java.awt.image.RenderedImage;
import java.io.IOException;

import org.zkoss.image.Image;
import org.zkoss.lang.Objects;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zk.ui.sys.ContentRenderer;
import org.zkoss.zkex.zul.Colorbox;
import org.zkoss.zul.impl.Utils;

/**
 * @author Sam
 *
 */
public class Colorbutton extends Colorbox implements org.zkoss.zss.app.zul.api.Colorbutton {
	private String _src;
	private Image _image;
	/** Count the version of {@link #_image}. */
	private byte _imgver;

	/** Returns the image URI.
	 * <p>Default: null.
	 */
	public String getImage() {
		return _src;
	}
	/** Sets the image URI.
	 * <p>Calling this method implies setImageContent(null).
	 * In other words, the last invocation of {@link #setImage} overrides
	 * the previous {@link #setImageContent}, if any.
	 * @see #setImageContent(Image)
	 * @see #setImageContent(RenderedImage)
	 */
	public void setImage(String src) {
		if (src != null && src.length() == 0) src = null;
		if (_image != null || !Objects.equals(_src, src)) {
			_src = src;
			_image = null;
			smartUpdate("image", new EncodedImageURL());
		}
	}

	public String getZclass() {
		return _zclass == null ? "z-colorbtn" : _zclass;
	}

	@Override
	protected void renderProperties(ContentRenderer renderer)
			throws IOException {
		super.renderProperties(renderer);

		render(renderer, "image", getEncodedImageURL());
	}

	/** Returns the encoded URL for the image ({@link #getImage}
	 * or {@link #getImageContent}), or null if no image.
	 * <p>Used only for component developements; not by app developers.
	 * <p>Note: this method can be invoked only if execution is not null.
	 */
	private String getEncodedImageURL() {
		if (_image != null)
			return Utils.getDynamicMediaURI( //already encoded
				this, _imgver, "c/" + _image.getName(), _image.getFormat());

		final Desktop dt = getDesktop(); //it might not belong to any desktop
		return dt != null && _src != null ?
			dt.getExecution().encodeURL(_src): null;
	}

	private class EncodedImageURL implements org.zkoss.zk.ui.util.DeferredValue {
		public Object getValue() {
			return getEncodedImageURL();
		}
	}
}
