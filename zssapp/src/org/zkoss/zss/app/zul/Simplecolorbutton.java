/* Colorbutton.java

{{IS_NOTE
	Purpose:

	Description:

	History:
		Apr 12, 2010 3:47:56 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.zul;

import java.io.IOException;
import java.util.Map;

import org.zkoss.image.Image;
import org.zkoss.lang.Objects;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.sys.ContentRenderer;
import org.zkoss.zul.impl.Utils;
import org.zkoss.zul.impl.XulElement;

/**
 * @author Sam
 *
 */
public class Simplecolorbutton extends XulElement implements org.zkoss.zss.app.zul.api.Colorbutton {
	private String _color = "#000000";
	private int _rgb = 0x000000;
	private boolean _noSmartUpdate = false;

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
	 */
	public void setImage(String src) {
		if (src != null && src.length() == 0) src = null;
		if (_image != null || !Objects.equals(_src, src)) {
			_src = src;
			_image = null;
			smartUpdate("image", new EncodedImageURL());
		}
	}
	/**
	 * Sets the color.
	 * @param color in #RRGGBB format (hexdecimal).
	 */
	public String getColor() {
		return _color;
	}
	/**
	 * Returns the color (in string as #RRGGBB).
	 * <p>Default: #000000
	 */
	public void setColor(String color) {
		if(!Objects.equals(color, _color)) {
			_color = color;
			_rgb = _color == null ? 0 : decode(_color);
			if (!_noSmartUpdate)
				smartUpdate("color", _color);
		}
	}
	/**
	 * Returns the color in int
	 * <p>Default: 0x000000
	 */
	public int getRGB() {
		return _rgb;
	}

	public String getZclass() {
		return _zclass == null ? "z-colorbtn" : _zclass;
	}

	@Override
	protected void renderProperties(ContentRenderer renderer)
			throws IOException {
		super.renderProperties(renderer);

		render(renderer, "image", getEncodedImageURL());
		renderer.render("color", getColor());
	}
	/** Processes an AU request. It is invoked internally.
	 *
	 * <p>Default: in addition to what are handled by {@link XulElement#service},
	 * it also handles onChange.
	 */
	public void service(AuRequest request, boolean everError) {
		final String cmd = request.getCommand();
		if (Events.ON_CHANGE.equals(cmd)) {
			final Map data = request.getData();

			_noSmartUpdate = true;
			setColor((String)data.get("color"));
			_noSmartUpdate = false;
			
			Events.postEvent(Event.getEvent(request));
		} else
			super.service(request, everError);
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

	private static int decode(String color) {
		if (color == null) {
			return 0;
		}
		if (color.length() != 7 || !color.startsWith("#")) {
			throw new UiException("Incorrect color format (#RRGGBB) : " + color);
		}
		return Integer.parseInt(color.substring(1), 16);
	}
}