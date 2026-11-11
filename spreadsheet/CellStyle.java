import java.awt.Color;

/**
 * Mutable formatting record for a single cell.
 * Stored by {@link SpreadsheetModel} in its style map and read by
 * {@link SpreadsheetCellRenderer} when painting.
 */
public class CellStyle {

    /** Alignment sentinel: numbers right-aligned, text left-aligned (Excel default). */
    public static final int ALIGN_AUTO = -1;

    public boolean bold;
    public boolean italic;
    /** One of SwingConstants.LEFT / CENTER / RIGHT, or {@link #ALIGN_AUTO}. */
    public int alignment = ALIGN_AUTO;
    /** Cell fill colour; null = default (white). */
    public Color background;
    /** Text colour; null = default. */
    public Color foreground;
    /** Font size in points; 0 = table default. */
    public int fontSize;

    public CellStyle copy() {
        CellStyle c = new CellStyle();
        c.bold = bold;
        c.italic = italic;
        c.alignment = alignment;
        c.background = background;
        c.foreground = foreground;
        c.fontSize = fontSize;
        return c;
    }

    /** Restores every attribute to its default value. */
    public void reset() {
        bold = false;
        italic = false;
        alignment = ALIGN_AUTO;
        background = null;
        foreground = null;
        fontSize = 0;
    }

    /** True when no attribute differs from the default (such styles are not stored). */
    public boolean isDefault() {
        return !bold && !italic && alignment == ALIGN_AUTO
                && background == null && foreground == null && fontSize == 0;
    }
}
