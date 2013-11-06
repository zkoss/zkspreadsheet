package org.zkoss.zss.test.selenium;

import java.awt.image.BufferedImage;

public interface Comparator {
	/**
	 * Return whether the two images are the same. If null, the images are the same,
	 * otherwise, the returned image is the change indicator.
	 * @param b1 the source of the base image
	 * @param b2 the compared image.
	 */
	public BufferedImage compare(BufferedImage b1, BufferedImage b2);
}