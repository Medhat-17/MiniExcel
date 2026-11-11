import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.ThreadLocalRandom;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JSpinner;
import javax.swing.SpinnerNumberModel;

/**
 * "Insert" tool: drops a date, random sample numbers, a sample sales table
 * or a chart placeholder into the sheet at the selected cell.
 */
public class InsertDialog extends JDialog {

    private static final Color GREEN = new Color(0x107C41);
    private static final String[] TYPES = {
            "Current date", "Current date and time", "Random sample numbers",
            "Sample sales table", "Chart placeholder"};

    private final SpreadsheetModel model;
    private final int startRow;
    private final int startCol;
    private final Runnable onDone;

    private final JComboBox<String> typeBox = new JComboBox<>(TYPES);
    private final JSpinner rowsSpin = new JSpinner(new SpinnerNumberModel(5, 1, 50, 1));
    private final JSpinner colsSpin = new JSpinner(new SpinnerNumberModel(4, 1, 26, 1));
    private final JSpinner minSpin = new JSpinner(new SpinnerNumberModel(1, -1000000, 1000000, 1));
    private final JSpinner maxSpin = new JSpinner(new SpinnerNumberModel(100, -1000000, 1000000, 1));

    public InsertDialog(JFrame owner, SpreadsheetModel model, int startRow, int startCol, Runnable onDone) {
        super(owner, "Insert", true);
        this.model = model;
        this.startRow = startRow;
        this.startCol = startCol;
        this.onDone = onDone;

        JPanel form = new JPanel(new GridLayout(0, 2, 8, 8));
        form.setBorder(BorderFactory.createEmptyBorder(14, 16, 8, 16));
        form.add(new JLabel("Insert at cell:"));
        JLabel where = new JLabel(FormulaEngine.cellRef(startRow, startCol));
        where.setForeground(GREEN);
        where.setFont(where.getFont().deriveFont(java.awt.Font.BOLD));
        form.add(where);
        form.add(new JLabel("What to insert:"));
        form.add(typeBox);
        form.add(new JLabel("Rows (random data):"));
        form.add(rowsSpin);
        form.add(new JLabel("Columns (random data):"));
        form.add(colsSpin);
        form.add(new JLabel("Minimum value:"));
        form.add(minSpin);
        form.add(new JLabel("Maximum value:"));
        form.add(maxSpin);

        typeBox.addActionListener(e -> updateEnabled());
        updateEnabled();

        JButton insert = new JButton("Insert");
        insert.addActionListener(e -> doInsert());
        JButton cancel = new JButton("Cancel");
        cancel.addActionListener(e -> dispose());
        JPanel buttons = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 8));
        buttons.add(insert);
        buttons.add(cancel);

        add(form, BorderLayout.CENTER);
        add(buttons, BorderLayout.SOUTH);
        getRootPane().setDefaultButton(insert);
        pack();
        setResizable(false);
        setLocationRelativeTo(owner);
    }

    private void updateEnabled() {
        boolean random = typeBox.getSelectedIndex() == 2;
        rowsSpin.setEnabled(random);
        colsSpin.setEnabled(random);
        minSpin.setEnabled(random);
        maxSpin.setEnabled(random);
    }

    private void doInsert() {
        model.beginBatch();
        try {
            switch (typeBox.getSelectedIndex()) {
                case 0:
                    model.setInput(startRow, startCol, LocalDate.now().toString());
                    break;
                case 1:
                    model.setInput(startRow, startCol,
                            LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm")));
                    break;
                case 2:
                    insertRandom();
                    break;
                case 3:
                    insertSalesTable();
                    break;
                default:
                    insertChartPlaceholder();
                    break;
            }
        } finally {
            model.endBatch();
        }
        dispose();
        if (onDone != null) onDone.run();
    }

    private void insertRandom() {
        int rows = (Integer) rowsSpin.getValue();
        int cols = (Integer) colsSpin.getValue();
        int lo = (Integer) minSpin.getValue();
        int hi = (Integer) maxSpin.getValue();
        if (hi < lo) {
            int t = lo;
            lo = hi;
            hi = t;
        }
        ThreadLocalRandom rnd = ThreadLocalRandom.current();
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                model.setInput(startRow + r, startCol + c, String.valueOf(lo + rnd.nextInt(hi - lo + 1)));
    }

    private void insertSalesTable() {
        String[] head = {"Month", "Sales", "Cost", "Profit"};
        String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun"};
        ThreadLocalRandom rnd = ThreadLocalRandom.current();
        for (int j = 0; j < head.length; j++) {
            model.setInput(startRow, startCol + j, head[j]);
            model.updateStyle(startRow, startCol + j, st -> {
                st.bold = true;
                st.background = GREEN;
                st.foreground = Color.WHITE;
            });
        }
        for (int i = 0; i < months.length; i++) {
            int r = startRow + 1 + i;
            model.setInput(r, startCol, months[i]);
            model.setInput(r, startCol + 1, String.valueOf(2000 + rnd.nextInt(3000)));
            model.setInput(r, startCol + 2, String.valueOf(800 + rnd.nextInt(1500)));
            model.setInput(r, startCol + 3,
                    "=" + FormulaEngine.cellRef(r, startCol + 1) + "-" + FormulaEngine.cellRef(r, startCol + 2));
        }
        int total = startRow + 1 + months.length;
        model.setInput(total, startCol, "Total");
        for (int j = 1; j < head.length; j++) {
            model.setInput(total, startCol + j, "=SUM(" + FormulaEngine.cellRef(startRow + 1, startCol + j)
                    + ":" + FormulaEngine.cellRef(total - 1, startCol + j) + ")");
        }
        for (int j = 0; j < head.length; j++) {
            model.updateStyle(total, startCol + j, st -> st.bold = true);
        }
    }

    /** A small framed block with sample data and text bars standing in for a real chart. */
    private void insertChartPlaceholder() {
        Color panel = new Color(0xEAF3EE);
        model.setInput(startRow, startCol, "[ CHART PLACEHOLDER ]");
        model.setInput(startRow + 1, startCol, "Category");
        model.setInput(startRow + 1, startCol + 1, "Value");
        model.setInput(startRow + 1, startCol + 2, "Bar");
        ThreadLocalRandom rnd = ThreadLocalRandom.current();
        for (int i = 0; i < 5; i++) {
            int r = startRow + 2 + i;
            int value = 10 + rnd.nextInt(81);
            StringBuilder bar = new StringBuilder();
            for (int k = 0; k < value / 10; k++) bar.append('\u2588');
            model.setInput(r, startCol, "Q" + (i + 1));
            model.setInput(r, startCol + 1, String.valueOf(value));
            model.setInput(r, startCol + 2, bar.toString());
            model.updateStyle(r, startCol + 2, st -> st.foreground = GREEN);
        }
        for (int r = startRow; r < startRow + 7; r++) {
            for (int c = startCol; c < startCol + 3; c++) {
                final boolean head = r <= startRow + 1;
                model.updateStyle(r, c, st -> {
                    st.background = panel;
                    if (head) st.bold = true;
                });
            }
        }
    }
}
