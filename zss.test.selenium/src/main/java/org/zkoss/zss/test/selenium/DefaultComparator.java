package org.zkoss.zss.test.selenium;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.image.BufferedImage;


/**
 * A comparer class to compare two images. (copied from ztl)
 * @author Phoenix
 * @author jumperchen
 */
public class DefaultComparator implements Comparator{
	private int _comparex, _comparey, _leniency;

    public DefaultComparator(int comparex, int comparey, int leniency) {
        this._comparex = comparex;
        this._comparey = comparey;
        this._leniency = leniency;
    }

    // buffered images are just better.
    private static BufferedImage imageToBufferedImage(Image img) {
        BufferedImage bi = new BufferedImage(img.getWidth(null), img.getHeight(null), BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = bi.createGraphics();
        g2.drawImage(img, null, null);
        return bi;
    }

    private int[][] getImageMap(BufferedImage img) {
        // setup brightness map
    	int width = img.getWidth();
        int height = img.getHeight();
        int[][] map = new int[height][width];
        int ta = 0;
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                ta = (int)(100 * getBrightnessAtPoint(img, x, y));
                map[y][x] = ta;
            }
        }
        return map;
    } 

    private static float getBrightnessAtPoint(BufferedImage img, int x, int y) {
        return getColorFactor(new Color(img.getRGB(x, y)));
    }

    private static float getColorFactor(Color c) {
        float[] hsb = Color.RGBtoHSB(c.getRed(), c.getGreen(), c.getBlue(), null);
        return (float) (hsb[2] * 0.5 + ((hsb[0] / 360) * 50/* 1/2 of 100 */));
    }

    /**
     * Compare 2 images with the same size
     * @param bi1
     * @param bi2
     * @return an image that shows the difference or null for 2 images are equivalent
     */
	public BufferedImage compare(BufferedImage bi1, BufferedImage bi2) {
	    if (bi1.getWidth() != bi2.getWidth() || bi1.getHeight() != bi2.getHeight()){
	        return bi2;
        }
    	int[][] m1 = getImageMap(bi1);
        int[][] m2 = getImageMap(bi2);
    	BufferedImage imgc = imageToBufferedImage(bi2);
        Graphics2D gc = imgc.createGraphics();
        gc.setColor(Color.RED);
        int cx = _comparex;
        int w = bi1.getWidth();
        int h = bi1.getHeight();
        if (cx > w) cx = w;
        int cy = _comparey;
        if (cy > h) cy = h;
        
        // how many points per section
        int bx = (int)(Math.floor(w / cx));
        if (bx <= 0) bx = 1;
        int by = (int)(Math.floor(h / cy));
        if (by <= 0) by = 1;
        int[][] variance = new int[cy][cx];
        
        // set to a match by default, if a change is found then flag non-match
        boolean match = true;
        // loop through whole image and compare individual blocks of images
        int ty = 0;
        for (int y = 0; y < cy; y++) {
            ty = y*by;
            for (int x = 0; x < cx; x++) {
                int b1 = aggregateMapArea(m1, x*bx, ty, bx, by);
                int b2 = aggregateMapArea(m2, x*bx, ty, bx, by);
                int diff = Math.abs(b1 - b2);
                variance[y][x] = diff; 
                if (diff > _leniency) { // the difference in a certain region has passed the threshold value
                    gc.drawRect(x*bx, y*by, bx - 1, by - 1);
                    match = false;
                }
            }
        }
        return match ? null : imgc;
    }
    private int aggregateMapArea(int[][] map, int ox, int oy, int w, int h) {
        int t = 0;
        for (int i = 0; i < h; i++) {
            int ty = oy+i;
            for (int j = 0; j < w; j++)
                t += map[ty][ox+j];
        }
        return (int)(t/(w*h));
    }

}