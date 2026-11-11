import java.awt.Graphics2D;
import java.awt.Rectangle;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import javax.imageio.ImageIO;

/**
 * A picture floating above the grid of one worksheet.
 *
 * The pixel data is kept as PNG bytes (so it can be saved inside a workbook file)
 * plus a decoded {@link BufferedImage} for painting. Position and size are stored
 * in table pixel coordinates and can be changed freely by the user.
 */
public final class SheetImage {

    /** Stored pixel data is downscaled to this width to keep workbook files small. */
    public static final int MAX_STORED_WIDTH = 1600;
    /** A freshly inserted picture is shown at most this big. */
    public static final int MAX_INITIAL_WIDTH = 480;
    public static final int MAX_INITIAL_HEIGHT = 360;

    public int x;
    public int y;
    public int width;
    public int height;

    private final byte[] png;
    private final BufferedImage image;

    private SheetImage(byte[] png, BufferedImage image, int x, int y, int width, int height) {
        this.png = png;
        this.image = image;
        this.x = x;
        this.y = y;
        this.width = width;
        this.height = height;
    }

    /** Creates a picture from any decoded image, placed with its top-left corner at (x, y). */
    public static SheetImage create(BufferedImage source, int x, int y) throws IOException {
        BufferedImage img = toArgb(source, MAX_STORED_WIDTH);
        double scale = Math.min(1.0, Math.min(
                (double) MAX_INITIAL_WIDTH / img.getWidth(),
                (double) MAX_INITIAL_HEIGHT / img.getHeight()));
        int w = Math.max(16, (int) Math.round(img.getWidth() * scale));
        int h = Math.max(16, (int) Math.round(img.getHeight() * scale));
        return new SheetImage(encode(img), img, x, y, w, h);
    }

    /** Restores a picture from saved PNG bytes. */
    public static SheetImage fromPng(byte[] data, int x, int y, int w, int h) throws IOException {
        BufferedImage img = ImageIO.read(new ByteArrayInputStream(data));
        if (img == null) throw new IOException("Unreadable picture data");
        return new SheetImage(data, img, x, y, Math.max(8, w), Math.max(8, h));
    }

    public BufferedImage getImage() {
        return image;
    }

    public byte[] getPngBytes() {
        return png;
    }

    public Rectangle getBounds() {
        return new Rectangle(x, y, width, height);
    }

    private static BufferedImage toArgb(BufferedImage src, int maxWidth) {
        int w = src.getWidth();
        int h = src.getHeight();
        if (w > maxWidth) {
            h = Math.max(1, (int) Math.round(h * (double) maxWidth / w));
            w = maxWidth;
        }
        BufferedImage out = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = out.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        g.drawImage(src, 0, 0, w, h, null);
        g.dispose();
        return out;
    }

    private static byte[] encode(BufferedImage img) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        if (!ImageIO.write(img, "png", bos)) throw new IOException("Could not encode picture");
        return bos.toByteArray();
    }
}
