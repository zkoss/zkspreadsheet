package org.zkoss.zss.test.selenium.util;

/**
 * To crop PNG bytes
 */
public interface PngCropper {
    byte[] crop(byte[] pngData);
}
