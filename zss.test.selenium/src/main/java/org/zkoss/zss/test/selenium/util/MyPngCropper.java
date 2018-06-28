package org.zkoss.zss.test.selenium.util;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;

/**
 * crop PNG with an offset of top, right, bottom, or left.
 */
public class MyPngCropper implements PngCropper {
    public static PngCropper TOOLBAR_CROPPER = new MyPngCropper(100, 0, 0, 0); //cut toolbar & formula bar on the top
    private final int left;
    private final int top;
    private final int right;
    private final int bottom;

    public MyPngCropper(int top, int right, int bottom, int left) {
        this.left = left;
        this.top = top;
        this.right = right;
        this.bottom = bottom;
    }

    @Override
    public byte[] crop(byte[] pngData) {
        try {
            BufferedImage image = ImageIO.read(new ByteArrayInputStream(pngData));
            BufferedImage subImage = image.getSubimage(left, top, image.getWidth() - left - right, image.getHeight() - top - bottom);
            ByteArrayOutputStream output = new ByteArrayOutputStream();
            ImageIO.write(subImage, "png", output);
            return output.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
