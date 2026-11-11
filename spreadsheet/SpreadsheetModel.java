import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;
import java.util.regex.Pattern;
import javax.swing.table.AbstractTableModel;

/**
 * Data-model layer: one worksheet grid (default 60 rows x 26 columns, headers A-Z).
 *
 * Three separate maps keyed by cell reference ("A1", "B7", ...):
 * <ul>
 *   <li>{@code rawValues} - literal text / numbers typed by the user</li>
 *   <li>{@code formulas}  - inputs that start with '='</li>
 *   <li>{@code styles}    - bold / italic / alignment / colours / font size</li>
 * </ul>
 * A fourth map caches computed formula results. Any edit clears the cache and
 * fires a table update, so every dependent formula is recalculated lazily the
 * next time it is painted or referenced (circular references yield #ERROR!).
 */
public class SpreadsheetModel extends AbstractTableModel {

    public static final int DEFAULT_ROWS = 60;
    public static final int DEFAULT_COLS = 26;

    private static final Pattern NUMBER =
            Pattern.compile("^[-+]?(\\d+(\\.\\d*)?|\\.\\d+)([eE][-+]?\\d+)?$");

    private final int rows;
    private final int cols;

    private final Map<String, String> rawValues = new HashMap<>();
    private final Map<String, String> formulas = new HashMap<>();
    private final Map<String, CellStyle> styles = new HashMap<>();
    private final Map<String, Object> computed = new HashMap<>();
    private final Set<String> evaluating = new HashSet<>();

    private final FormulaEngine engine;

    /** Sheet protection state and the pictures floating above the grid. */
    private final SheetProtection protection = new SheetProtection();
    private final List<SheetImage> images = new ArrayList<>();

    private int batchDepth = 0;
    private boolean dirty = false;

    public SpreadsheetModel() {
        this(DEFAULT_ROWS, DEFAULT_COLS);
    }

    public SpreadsheetModel(int rows, int cols) {
        this.rows = rows;
        this.cols = cols;
        this.engine = new FormulaEngine(this::getComputedValue, rows, cols);
    }

    // ------------------------------------------------------------------
    // TableModel contract
    // ------------------------------------------------------------------

    @Override
    public int getRowCount() {
        return rows;
    }

    @Override
    public int getColumnCount() {
        return cols;
    }

    @Override
    public String getColumnName(int column) {
        return FormulaEngine.columnLabel(column);
    }

    @Override
    public Class<?> getColumnClass(int columnIndex) {
        return Object.class;
    }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        return !protection.isProtected();
    }

    // ------------------------------------------------------------------
    // Protection and pictures
    // ------------------------------------------------------------------

    public SheetProtection getProtection() {
        return protection;
    }

    public boolean isProtected() {
        return protection.isProtected();
    }

    /** Read-only view of the pictures, back to front. */
    public List<SheetImage> getImages() {
        return Collections.unmodifiableList(images);
    }

    /** Adds a picture (ignored while the sheet is protected). */
    public void addImage(SheetImage image) {
        if (!protection.isProtected()) images.add(image);
    }

    public void removeImage(SheetImage image) {
        if (!protection.isProtected()) images.remove(image);
    }

    /** Display text for the cell: formula results are formatted, empty cells give "". */
    @Override
    public Object getValueAt(int row, int col) {
        Object v = getComputedValue(row, col);
        if (v == null) return "";
        if (v instanceof Double) return FormulaEngine.formatNumber((Double) v);
        return v.toString();
    }

    @Override
    public void setValueAt(Object value, int row, int col) {
        setInput(row, col, value == null ? "" : value.toString());
    }

    // ------------------------------------------------------------------
    // Cell content
    // ------------------------------------------------------------------

    public boolean inBounds(int row, int col) {
        return row >= 0 && col >= 0 && row < rows && col < cols;
    }

    private static String key(int row, int col) {
        return FormulaEngine.cellRef(row, col);
    }

    /** What the user typed: the formula text (with '=') or the literal value. */
    public String getInput(int row, int col) {
        if (!inBounds(row, col)) return "";
        String k = key(row, col);
        String f = formulas.get(k);
        if (f != null) return f;
        String r = rawValues.get(k);
        return r == null ? "" : r;
    }

    /** Stores user input, routing it to the formula or raw-value map, then recalculates. */
    public void setInput(int row, int col, String text) {
        if (!inBounds(row, col) || protection.isProtected()) return;
        String k = key(row, col);
        String t = text == null ? "" : text.trim();
        formulas.remove(k);
        rawValues.remove(k);
        if (t.startsWith("=")) {
            formulas.put(k, t);
        } else if (!t.isEmpty()) {
            rawValues.put(k, t);
        }
        recalculate();
    }

    /** Removes the content of a cell but keeps its formatting (like pressing Delete). */
    public void clearContents(int row, int col) {
        setInput(row, col, "");
    }

    /**
     * Evaluated value: Double, String (text or error code) or null (empty).
     * Formula results are memoised in the cache until the next edit.
     */
    public Object getComputedValue(int row, int col) {
        if (!inBounds(row, col)) return null;
        String k = key(row, col);
        String f = formulas.get(k);
        if (f != null) {
            if (computed.containsKey(k)) return computed.get(k);
            if (!evaluating.add(k)) return FormulaEngine.ERR_GENERIC; // circular reference
            Object result;
            try {
                result = engine.evaluate(f);
            } finally {
                evaluating.remove(k);
            }
            computed.put(k, result);
            return result;
        }
        String raw = rawValues.get(k);
        if (raw == null) return null;
        Double d = tryParseNumber(raw);
        return d != null ? (Object) d : raw;
    }

    private static Double tryParseNumber(String s) {
        String t = s.trim();
        if (!NUMBER.matcher(t).matches()) return null;
        try {
            return Double.valueOf(t);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    public boolean isNumeric(int row, int col) {
        return getComputedValue(row, col) instanceof Double;
    }

    public boolean isFormulaError(int row, int col) {
        if (!inBounds(row, col) || !formulas.containsKey(key(row, col))) return false;
        Object v = getComputedValue(row, col);
        return v instanceof String && FormulaEngine.isErrorCode((String) v);
    }

    // ------------------------------------------------------------------
    // Formatting
    // ------------------------------------------------------------------

    /** @return the cell's style or null when it has default formatting. */
    public CellStyle getStyle(int row, int col) {
        return inBounds(row, col) ? styles.get(key(row, col)) : null;
    }

    /** Applies a mutation to the cell's style (created on demand) and repaints it. */
    public void updateStyle(int row, int col, Consumer<CellStyle> change) {
        if (!inBounds(row, col) || protection.isProtected()) return;
        String k = key(row, col);
        CellStyle st = styles.get(k);
        if (st == null) st = new CellStyle();
        change.accept(st);
        if (st.isDefault()) styles.remove(k);
        else styles.put(k, st);
        if (batchDepth == 0) fireTableCellUpdated(row, col);
        else dirty = true;
    }

    // ------------------------------------------------------------------
    // Recalculation / batching
    // ------------------------------------------------------------------

    /** Invalidates cached results; dependent formulas recompute on next access. */
    private void recalculate() {
        computed.clear();
        if (batchDepth == 0) fireTableRowsUpdated(0, rows - 1);
        else dirty = true;
    }

    /** Suppresses repaint events until the matching {@link #endBatch()} (used by macros / inserts). */
    public void beginBatch() {
        batchDepth++;
    }

    public void endBatch() {
        if (batchDepth > 0 && --batchDepth == 0 && dirty) {
            dirty = false;
            computed.clear();
            fireTableRowsUpdated(0, rows - 1);
        }
    }

    // ------------------------------------------------------------------
    // Used range (for CSV export)
    // ------------------------------------------------------------------

    public int getUsedRowCount() {
        int last = 0;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                if (!getInput(r, c).isEmpty()) last = r + 1;
        return last;
    }

    public int getUsedColumnCount() {
        int last = 0;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                if (!getInput(r, c).isEmpty() && c + 1 > last) last = c + 1;
        return last;
    }
}
