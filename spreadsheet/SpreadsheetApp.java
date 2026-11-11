import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Cursor;
import java.awt.BasicStroke;
import java.awt.Rectangle;
import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.image.BufferedImage;
import java.util.Arrays;
import javax.imageio.ImageIO;
import javax.swing.Icon;
import javax.swing.JPasswordField;
import javax.swing.event.AncestorEvent;
import javax.swing.event.AncestorListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.event.ActionEvent;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import javax.swing.AbstractAction;
import javax.swing.BorderFactory;
import javax.swing.ButtonModel;
import javax.swing.DefaultCellEditor;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.CompoundBorder;
import javax.swing.border.EmptyBorder;
import javax.swing.border.MatteBorder;
import javax.swing.event.ListSelectionListener;
import javax.swing.plaf.basic.BasicButtonUI;

/**
 * UI &amp; layout layer: the main window.
 *
 * <ul>
 *   <li>Excel-style green ribbon (#107C41) with all tools</li>
 *   <li>Formula bar: cell badge, "fx" label and a formula text field synced with the selection</li>
 *   <li>Multi-sheet workbook in a JTabbedPane (add / rename / delete tabs)</li>
 *   <li>Status bar with Count / Sum / Average of the selected cells</li>
 * </ul>
 * Each worksheet is a {@link SpreadsheetModel} shown in a JTable that uses
 * {@link SpreadsheetCellRenderer}; formulas are evaluated by {@link FormulaEngine}.
 */
public class SpreadsheetApp extends JFrame {

    static final Color GREEN = new Color(0x107C41);
    static final Color LIGHT_GRAY = new Color(0xF3F3F3);
    static final Color GRID = new Color(0xD0D0D0);

    private final JTabbedPane tabs = new JTabbedPane();
    private final JLabel cellBadge = new JLabel("A1", SwingConstants.CENTER);
    private final JTextField formulaField = new JTextField();
    private final JLabel statusLabel = new JLabel("Ready");
    private final JLabel statsLabel = new JLabel("");
    private final JComboBox<Integer> sizeCombo = new JComboBox<>(new Integer[]{9, 10, 11, 12, 13, 14, 16, 18, 20});
    private final JPanel fillChip = new JPanel();

    private Color currentBackground = new Color(0xFFF2CC);
    private int sheetCounter = 0;
    private boolean updatingUi = false;
    private MacroIDEDialog macroDialog;

    private final JLabel protLabel = new JLabel("", SwingConstants.CENTER);
    private JButton protectButton;
    private File currentFile;
    private char[] workbookPassword;
    private boolean dirty = false;
    private boolean loading = false;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception ignored) {
                // fall back to the default look and feel
            }
            new SpreadsheetApp().setVisible(true);
        });
    }

    public SpreadsheetApp() {
        super("Spreadsheet Pro");
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                exitApp();
            }
        });
        setSize(1360, 780);
        setLocationRelativeTo(null);
        setJMenuBar(buildMenuBar());

        JPanel north = new JPanel(new BorderLayout());
        north.add(buildRibbon(), BorderLayout.NORTH);
        north.add(buildFormulaBar(), BorderLayout.SOUTH);
        add(north, BorderLayout.NORTH);

        tabs.setTabPlacement(JTabbedPane.BOTTOM);
        tabs.addChangeListener(e -> syncFormulaBar());
        add(tabs, BorderLayout.CENTER);
        add(buildStatusBar(), BorderLayout.SOUTH);

        addSheet(nextSheetName());
        dirty = false;
        updateTitle();
    }

    // ==================================================================
    // Ribbon
    // ==================================================================

    private JComponent buildRibbon() {
        JPanel ribbon = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 6));
        ribbon.setBackground(GREEN);

        ribbon.add(button("Save", "Save the workbook (Ctrl+S)", e -> saveWorkbook()));
        ribbon.add(button("New Sheet", "Add a new worksheet (Shift+F11)", e -> addSheet(nextSheetName())));
        ribbon.add(button("Delete Sheet", "Delete the current worksheet", e -> deleteSheet(tabs.getSelectedIndex())));
        ribbon.add(separator());
        ribbon.add(button("\u03A3 AutoSum", "Sum the numbers above or to the left", e -> autoSum()));
        ribbon.add(button("Insert", "Insert date, sample data or a chart placeholder", e -> openInsertDialog()));
        ribbon.add(button("Picture", "Insert a picture from a file", e -> insertPictureFromFile()));
        ribbon.add(button("VBA Macro", "Open the macro editor and runner", e -> openMacroIde()));
        protectButton = button("Protect Sheet", "Protect or unprotect this sheet, optionally with a password",
                e -> toggleProtection());
        ribbon.add(protectButton);
        ribbon.add(separator());

        JButton bold = button("B", "Bold (Ctrl+B)", e -> toggleStyle(true));
        bold.setFont(bold.getFont().deriveFont(Font.BOLD, 14f));
        JButton italic = button("I", "Italic (Ctrl+I)", e -> toggleStyle(false));
        italic.setFont(italic.getFont().deriveFont(Font.BOLD | Font.ITALIC, 14f));
        ribbon.add(bold);
        ribbon.add(italic);
        ribbon.add(button("Left", "Align left (Ctrl+L)", e -> align(SwingConstants.LEFT)));
        ribbon.add(button("Center", "Align center (Ctrl+E)", e -> align(SwingConstants.CENTER)));
        ribbon.add(button("Right", "Align right (Ctrl+R)", e -> align(SwingConstants.RIGHT)));
        ribbon.add(separator());

        JLabel sizeLabel = new JLabel("Size");
        sizeLabel.setForeground(Color.WHITE);
        ribbon.add(sizeLabel);
        sizeCombo.setSelectedItem(13);
        sizeCombo.setFocusable(false);
        sizeCombo.setToolTipText("Font size of the selected cells");
        sizeCombo.addActionListener(e -> {
            if (updatingUi || sizeCombo.getSelectedItem() == null) return;
            final int size = (Integer) sizeCombo.getSelectedItem();
            applyToSelection(st -> st.fontSize = size);
        });
        ribbon.add(sizeCombo);
        ribbon.add(separator());

        ribbon.add(button("Fill Color", "Pick a background colour for the selected cells", e -> pickFillColor()));
        fillChip.setPreferredSize(new Dimension(18, 18));
        fillChip.setBackground(currentBackground);
        fillChip.setBorder(BorderFactory.createLineBorder(Color.WHITE, 1));
        ribbon.add(fillChip);
        ribbon.add(button("No Fill", "Remove the background colour", e -> applyToSelection(st -> st.background = null)));
        return ribbon;
    }

    private JButton button(String text, String tip, java.awt.event.ActionListener action) {
        RibbonButton b = new RibbonButton(text, tip);
        b.addActionListener(action);
        return b;
    }

    private JComponent separator() {
        JPanel p = new JPanel();
        p.setPreferredSize(new Dimension(1, 26));
        p.setBackground(new Color(0x4FAE7B));
        return p;
    }

    /** Flat ribbon button with a translucent hover / pressed highlight. */
    private static final class RibbonButton extends JButton {
        RibbonButton(String text, String tip) {
            super(text);
            setUI(new BasicButtonUI());
            setRolloverEnabled(true);
            setToolTipText(tip);
            setForeground(Color.WHITE);
            setFont(new Font(Font.SANS_SERIF, Font.BOLD, 12));
            setFocusPainted(false);
            setFocusable(false); // keeps keyboard focus in the grid
            setBorderPainted(false);
            setContentAreaFilled(false);
            setOpaque(false);
            setBorder(new EmptyBorder(6, 10, 6, 10));
            setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        }

        @Override
        protected void paintComponent(Graphics g) {
            ButtonModel m = getModel();
            if (m.isPressed() || m.isRollover()) {
                Graphics2D g2 = (Graphics2D) g.create();
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g2.setColor(m.isPressed() ? new Color(0, 0, 0, 60) : new Color(255, 255, 255, 45));
                g2.fillRoundRect(0, 0, getWidth(), getHeight(), 8, 8);
                g2.dispose();
            }
            super.paintComponent(g);
        }
    }

    // ==================================================================
    // Formula bar and status bar
    // ==================================================================

    private JComponent buildFormulaBar() {
        JPanel bar = new JPanel(new BorderLayout(6, 0));
        bar.setBackground(Color.WHITE);
        bar.setBorder(new CompoundBorder(new MatteBorder(0, 0, 1, 0, GRID), new EmptyBorder(4, 6, 4, 6)));

        cellBadge.setPreferredSize(new Dimension(90, 26));
        cellBadge.setOpaque(true);
        cellBadge.setBackground(LIGHT_GRAY);
        cellBadge.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 13));
        cellBadge.setBorder(BorderFactory.createLineBorder(GRID));

        JLabel fx = new JLabel("fx");
        fx.setFont(new Font(Font.SERIF, Font.BOLD | Font.ITALIC, 16));
        fx.setForeground(GREEN);
        fx.setBorder(new EmptyBorder(0, 6, 0, 6));

        JPanel left = new JPanel(new BorderLayout(6, 0));
        left.setOpaque(false);
        left.add(cellBadge, BorderLayout.WEST);
        left.add(fx, BorderLayout.CENTER);

        formulaField.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 13));
        formulaField.setBorder(new CompoundBorder(BorderFactory.createLineBorder(GRID), new EmptyBorder(2, 6, 2, 6)));
        formulaField.addActionListener(e -> commitFormulaBar());
        formulaField.getInputMap().put(KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), "revert");
        formulaField.getActionMap().put("revert", new AbstractAction() {
            @Override
            public void actionPerformed(ActionEvent e) {
                syncFormulaBar();
                SheetView sv = currentSheet();
                if (sv != null) sv.table.requestFocusInWindow();
            }
        });

        bar.add(left, BorderLayout.WEST);
        bar.add(formulaField, BorderLayout.CENTER);
        return bar;
    }

    private JComponent buildStatusBar() {
        JPanel bar = new JPanel(new BorderLayout());
        bar.setBackground(LIGHT_GRAY);
        bar.setBorder(new CompoundBorder(new MatteBorder(1, 0, 0, 0, GRID), new EmptyBorder(3, 10, 3, 10)));
        bar.add(statusLabel, BorderLayout.WEST);
        protLabel.setForeground(new Color(0xC00000));
        protLabel.setFont(protLabel.getFont().deriveFont(Font.BOLD));
        bar.add(protLabel, BorderLayout.CENTER);
        bar.add(statsLabel, BorderLayout.EAST);
        return bar;
    }

    private void status(String message) {
        statusLabel.setText(message);
    }

    /** Pushes the current selection into the badge, formula field, size box and stats label. */
    private void syncFormulaBar() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        updateProtectUi(sv);
        JTable t = sv.table;
        int r = t.getSelectionModel().getLeadSelectionIndex();
        int c = t.getColumnModel().getSelectionModel().getLeadSelectionIndex();
        if (!sv.model.inBounds(r, c)) {
            cellBadge.setText("");
            formulaField.setText("");
            statsLabel.setText("");
            return;
        }
        int[] rows = t.getSelectedRows();
        int[] cols = t.getSelectedColumns();
        if (rows.length * cols.length > 1) {
            cellBadge.setText(FormulaEngine.cellRef(rows[0], cols[0]) + ":"
                    + FormulaEngine.cellRef(rows[rows.length - 1], cols[cols.length - 1]));
        } else {
            cellBadge.setText(FormulaEngine.cellRef(r, c));
        }
        formulaField.setText(sv.model.getInput(r, c));
        formulaField.setCaretPosition(0);

        CellStyle st = sv.model.getStyle(r, c);
        updatingUi = true;
        sizeCombo.setSelectedItem(st != null && st.fontSize > 0 ? st.fontSize : 13);
        updatingUi = false;

        // Excel-like statistics for multi-cell selections
        int count = 0;
        double sum = 0;
        if (rows.length * cols.length > 1) {
            for (int rr : rows) {
                for (int cc : cols) {
                    Object v = sv.model.getComputedValue(rr, cc);
                    if (v instanceof Double) {
                        count++;
                        sum += (Double) v;
                    }
                }
            }
        }
        statsLabel.setText(count > 0
                ? "Count: " + count + "     Sum: " + FormulaEngine.formatNumber(sum)
                + "     Average: " + FormulaEngine.formatNumber(sum / count)
                : "");
    }

    /** Enter in the formula bar: write the text into the active cell and move down. */
    private void commitFormulaBar() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        int r = sv.table.getSelectionModel().getLeadSelectionIndex();
        int c = sv.table.getColumnModel().getSelectionModel().getLeadSelectionIndex();
        if (!sv.model.inBounds(r, c)) return;
        if (!guardEditable(sv)) return;
        if (sv.table.isEditing()) sv.table.getCellEditor().cancelCellEditing();
        sv.model.setInput(r, c, formulaField.getText());
        if (r + 1 < sv.model.getRowCount()) sv.table.changeSelection(r + 1, c, false, false);
        sv.table.requestFocusInWindow();
    }

    // ==================================================================
    // Menu bar
    // ==================================================================

    private JMenuItem menuItem(String text, int key, int modifiers, Runnable action) {
        JMenuItem item = new JMenuItem(text);
        if (key != 0) item.setAccelerator(KeyStroke.getKeyStroke(key, modifiers));
        item.addActionListener(e -> action.run());
        return item;
    }

    private JMenuBar buildMenuBar() {
        int ctrl = InputEvent.CTRL_DOWN_MASK;
        JMenuBar bar = new JMenuBar();

        JMenu file = new JMenu("File");
        file.add(menuItem("New Workbook", KeyEvent.VK_N, ctrl, this::newWorkbook));
        file.add(menuItem("Open...", KeyEvent.VK_O, ctrl, this::openWorkbook));
        file.add(menuItem("Save", KeyEvent.VK_S, ctrl, () -> saveWorkbook()));
        file.add(menuItem("Save As...", KeyEvent.VK_S, ctrl | InputEvent.SHIFT_DOWN_MASK, () -> saveWorkbookAs()));
        file.addSeparator();
        file.add(menuItem("Set Open Password...", 0, 0, this::setOpenPassword));
        file.addSeparator();
        file.add(menuItem("Import CSV...", 0, 0, this::importCsv));
        file.add(menuItem("Export CSV...", 0, 0, this::exportCsv));
        file.addSeparator();
        file.add(menuItem("Exit", 0, 0, this::exitApp));

        JMenu edit = new JMenu("Edit");
        edit.add(menuItem("Clear Contents", 0, 0, this::clearSelection));
        edit.add(menuItem("Clear Formatting", 0, 0, () -> applyToSelection(CellStyle::reset)));
        edit.addSeparator();
        edit.add(menuItem("Delete Selected Picture", 0, 0, this::deleteSelectedPicture));

        JMenu insert = new JMenu("Insert");
        insert.add(menuItem("New Sheet", KeyEvent.VK_F11, InputEvent.SHIFT_DOWN_MASK, () -> addSheet(nextSheetName())));
        insert.addSeparator();
        insert.add(menuItem("Picture from File...", 0, 0, this::insertPictureFromFile));
        insert.add(menuItem("Picture from Clipboard", 0, 0, this::insertPictureFromClipboard));
        insert.addSeparator();
        insert.add(menuItem("Date / Sample Data / Chart...", 0, 0, this::openInsertDialog));

        JMenu review = new JMenu("Review");
        review.add(menuItem("Protect Sheet...", 0, 0, this::protectSheetDialog));
        review.add(menuItem("Unprotect Sheet...", 0, 0, this::unprotectSheetDialog));

        JMenu help = new JMenu("Help");
        help.add(menuItem("Formula Reference", 0, 0, this::showFormulaHelp));

        bar.add(file);
        bar.add(edit);
        bar.add(insert);
        bar.add(review);
        bar.add(help);
        return bar;
    }

    private void showFormulaHelp() {
        JTextArea text = new JTextArea(
                "Formulas start with '='.\n\n"
                + "Operators:   +  -  *  /  ^  %   and parentheses\n"
                + "References:  =A1+B2*5\n"
                + "Ranges:      =SUM(A1:A10)\n\n"
                + "Functions:\n"
                + "  SUM, AVERAGE, MIN, MAX, COUNT  (ranges and lists: =MAX(A1:A5, 10))\n"
                + "  ABS, SQRT, ROUND(x, digits)\n\n"
                + "Errors:\n"
                + "  #DIV/0!   division by zero\n"
                + "  #SYNTAX!  malformed formula\n"
                + "  #NAME?    unknown function or name\n"
                + "  #ERROR!   bad reference, text used as a number, circular reference\n\n"
                + "Shortcuts: Ctrl+B bold, Ctrl+I italic, Ctrl+L/E/R align, Delete clear, F2 edit cell");
        text.setEditable(false);
        text.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 13));
        JOptionPane.showMessageDialog(this, text, "Formula Reference", JOptionPane.INFORMATION_MESSAGE);
    }

    // ==================================================================
    // Sheets (tabs)
    // ==================================================================

    private String nextSheetName() {
        return "Sheet" + (++sheetCounter);
    }

    private SheetView currentSheet() {
        Component c = tabs.getSelectedComponent();
        return c instanceof SheetView ? (SheetView) c : null;
    }

    private SheetView addSheet(String name) {
        return addSheet(name, new SpreadsheetModel());
    }

    private SheetView addSheet(String name, SpreadsheetModel model) {
        SheetView sv = new SheetView(model);
        tabs.addTab(name, sv);
        int idx = tabs.getTabCount() - 1;
        TabHeader header = new TabHeader(name);
        header.setLocked(model.isProtected());
        tabs.setTabComponentAt(idx, header);
        tabs.setSelectedIndex(idx);
        SwingUtilities.invokeLater(() -> sv.table.requestFocusInWindow());
        status("Added " + name);
        markDirty();
        return sv;
    }

    private void deleteSheet(int index) {
        if (index < 0) return;
        if (tabs.getTabCount() <= 1) {
            JOptionPane.showMessageDialog(this, "A workbook must contain at least one sheet.",
                    "Delete Sheet", JOptionPane.WARNING_MESSAGE);
            return;
        }
        Component comp = tabs.getComponentAt(index);
        if (comp instanceof SheetView && ((SheetView) comp).model.isProtected()) {
            JOptionPane.showMessageDialog(this, "This sheet is protected. Unprotect it before deleting it.",
                    "Delete Sheet", JOptionPane.WARNING_MESSAGE);
            return;
        }
        String title = tabs.getTitleAt(index);
        int answer = JOptionPane.showConfirmDialog(this,
                "Delete \"" + title + "\"? This cannot be undone.", "Delete Sheet",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.WARNING_MESSAGE);
        if (answer == JOptionPane.OK_OPTION) {
            tabs.removeTabAt(index);
            status("Deleted " + title);
            markDirty();
        }
    }

    private void renameSheet(int index) {
        if (index < 0) return;
        Object input = JOptionPane.showInputDialog(this, "Sheet name:", "Rename Sheet",
                JOptionPane.PLAIN_MESSAGE, null, null, tabs.getTitleAt(index));
        if (input == null) return;
        String name = input.toString().trim();
        if (name.isEmpty()) return;
        tabs.setTitleAt(index, name);
        Component header = tabs.getTabComponentAt(index);
        if (header instanceof TabHeader) ((TabHeader) header).setTitle(name);
        markDirty();
    }

    /** Custom tab title with a close button, a rename-on-double-click and a context menu. */
    private final class TabHeader extends JPanel {
        private final JLabel title;
        private final JLabel lockLabel = new JLabel(new LockIcon());

        TabHeader(String name) {
            super(new FlowLayout(FlowLayout.LEFT, 4, 0));
            setOpaque(false);
            lockLabel.setVisible(false);
            lockLabel.setToolTipText("This sheet is protected");
            add(lockLabel);
            title = new JLabel(name);
            add(title);

            JButton close = new JButton("\u00D7");
            close.setFocusable(false);
            close.setBorder(new EmptyBorder(0, 3, 0, 3));
            close.setContentAreaFilled(false);
            close.setFont(close.getFont().deriveFont(Font.BOLD, 14f));
            close.setForeground(Color.GRAY);
            close.setToolTipText("Close sheet");
            close.addActionListener(e -> deleteSheet(tabs.indexOfTabComponent(TabHeader.this)));
            close.addMouseListener(new MouseAdapter() {
                @Override
                public void mouseEntered(MouseEvent e) {
                    close.setForeground(new Color(0xC00000));
                }

                @Override
                public void mouseExited(MouseEvent e) {
                    close.setForeground(Color.GRAY);
                }
            });
            add(close);

            MouseAdapter ma = new MouseAdapter() {
                @Override
                public void mousePressed(MouseEvent e) {
                    int i = tabs.indexOfTabComponent(TabHeader.this);
                    if (i >= 0) tabs.setSelectedIndex(i);
                    if (e.isPopupTrigger()) showMenu(e);
                }

                @Override
                public void mouseReleased(MouseEvent e) {
                    if (e.isPopupTrigger()) showMenu(e);
                }

                @Override
                public void mouseClicked(MouseEvent e) {
                    if (e.getClickCount() == 2 && SwingUtilities.isLeftMouseButton(e)) {
                        renameSheet(tabs.indexOfTabComponent(TabHeader.this));
                    }
                }
            };
            addMouseListener(ma);
            title.addMouseListener(ma);
        }

        void setTitle(String name) {
            title.setText(name);
        }

        void setLocked(boolean locked) {
            lockLabel.setVisible(locked);
            revalidate();
            repaint();
        }

        private void showMenu(MouseEvent e) {
            JPopupMenu menu = new JPopupMenu();
            JMenuItem rename = new JMenuItem("Rename...");
            rename.addActionListener(a -> renameSheet(tabs.indexOfTabComponent(TabHeader.this)));
            JMenuItem delete = new JMenuItem("Delete");
            delete.addActionListener(a -> deleteSheet(tabs.indexOfTabComponent(TabHeader.this)));
            menu.add(rename);
            menu.add(delete);
            menu.show(e.getComponent(), e.getX(), e.getY());
        }
    }

    // ==================================================================
    // One worksheet view (table + headers)
    // ==================================================================

    private final class SheetView extends JPanel {
        final SpreadsheetModel model;
        final JTable table;
        final JList<String> rowHeader;
        ImageLayer layer; // pictures floating above the grid

        SheetView(SpreadsheetModel sheetModel) {
            super(new BorderLayout());
            this.model = sheetModel;

            table = new JTable(model) {
                @Override
                protected void paintComponent(Graphics g) {
                    super.paintComponent(g);
                    if (SheetView.this.layer != null) SheetView.this.layer.paint(g);
                }

                @Override
                protected void processMouseEvent(MouseEvent e) {
                    if (SheetView.this.layer != null && SheetView.this.layer.handleMouse(e)) return;
                    super.processMouseEvent(e);
                }

                @Override
                protected void processMouseMotionEvent(MouseEvent e) {
                    if (SheetView.this.layer != null && SheetView.this.layer.handleMotion(e)) return;
                    super.processMouseMotionEvent(e);
                }

                @Override
                public boolean editCellAt(int row, int column, java.util.EventObject e) {
                    if (SheetView.this.model.isProtected()) {
                        guardEditable(SheetView.this);
                        return false;
                    }
                    return super.editCellAt(row, column, e);
                }

                /** Typing over a selected cell replaces its content, like Excel. */
                @Override
                protected boolean processKeyBinding(KeyStroke ks, KeyEvent e, int condition, boolean pressed) {
                    boolean wasEditing = isEditing();
                    boolean handled = super.processKeyBinding(ks, e, condition, pressed);
                    if (!wasEditing && isEditing() && e.getID() == KeyEvent.KEY_PRESSED
                            && e.getKeyChar() != KeyEvent.CHAR_UNDEFINED
                            && getEditorComponent() instanceof JTextField) {
                        ((JTextField) getEditorComponent()).selectAll();
                    }
                    return handled;
                }
            };
            table.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 13));
            table.setRowHeight(22);
            table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
            table.setCellSelectionEnabled(true);
            table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
            table.setShowGrid(false);
            table.setIntercellSpacing(new Dimension(0, 0));
            table.getTableHeader().setReorderingAllowed(false);
            table.getTableHeader().setDefaultRenderer(new SpreadsheetCellRenderer.ColumnHeaderRenderer());
            table.getTableHeader().setPreferredSize(new Dimension(100, 24));
            for (int c = 0; c < model.getColumnCount(); c++) {
                table.getColumnModel().getColumn(c).setPreferredWidth(90);
            }
            table.setDefaultRenderer(Object.class, new SpreadsheetCellRenderer(model));
            table.setDefaultEditor(Object.class, new FormulaCellEditor(model, table.getFont()));
            layer = new ImageLayer(table, model, SpreadsheetApp.this::markDirty);

            String[] labels = new String[model.getRowCount()];
            for (int i = 0; i < labels.length; i++) labels[i] = String.valueOf(i + 1);
            rowHeader = new JList<>(labels);
            rowHeader.setFixedCellHeight(table.getRowHeight());
            rowHeader.setFixedCellWidth(46);
            rowHeader.setCellRenderer(new SpreadsheetCellRenderer.RowHeaderRenderer(table));
            rowHeader.setBackground(SpreadsheetCellRenderer.HEADER_BG);
            rowHeader.addMouseListener(new MouseAdapter() {
                @Override
                public void mousePressed(MouseEvent e) {
                    int i = rowHeader.locationToIndex(e.getPoint());
                    if (i < 0) return;
                    table.setRowSelectionInterval(i, i);
                    table.setColumnSelectionInterval(model.getColumnCount() - 1, 0); // active cell = column A
                    table.requestFocusInWindow();
                }
            });

            JScrollPane scroll = new JScrollPane(table);
            scroll.setRowHeaderView(rowHeader);
            JPanel corner = new JPanel();
            corner.setBackground(SpreadsheetCellRenderer.HEADER_BG);
            scroll.setCorner(JScrollPane.UPPER_LEFT_CORNER, corner);
            scroll.setBorder(BorderFactory.createEmptyBorder());
            scroll.getVerticalScrollBar().setUnitIncrement(22);
            add(scroll, BorderLayout.CENTER);

            installKeyBindings();

            ListSelectionListener l = e -> {
                if (e.getValueIsAdjusting()) return;
                if (SheetView.this == currentSheet()) syncFormulaBar();
                rowHeader.repaint();
                table.getTableHeader().repaint();
            };
            table.getSelectionModel().addListSelectionListener(l);
            table.getColumnModel().getSelectionModel().addListSelectionListener(l);

            // Reactive formula bar: whenever the model changes (typing, macros, insert) refresh it
            model.addTableModelListener(e -> {
                markDirty();
                if (SheetView.this == currentSheet()) syncFormulaBar();
            });

            table.changeSelection(0, 0, false, false);
        }

        private void installKeyBindings() {
            bind(KeyEvent.VK_B, InputEvent.CTRL_DOWN_MASK, "bold", () -> toggleStyle(true));
            bind(KeyEvent.VK_I, InputEvent.CTRL_DOWN_MASK, "italic", () -> toggleStyle(false));
            bind(KeyEvent.VK_L, InputEvent.CTRL_DOWN_MASK, "alignLeft", () -> align(SwingConstants.LEFT));
            bind(KeyEvent.VK_E, InputEvent.CTRL_DOWN_MASK, "alignCenter", () -> align(SwingConstants.CENTER));
            bind(KeyEvent.VK_R, InputEvent.CTRL_DOWN_MASK, "alignRight", () -> align(SwingConstants.RIGHT));
            bind(KeyEvent.VK_DELETE, 0, "clearContents", () -> {
                if (layer.hasSelection()) layer.deleteSelected();
                else clearSelection();
            });
        }

        private void bind(int key, int modifiers, String name, Runnable action) {
            table.getInputMap(JComponent.WHEN_FOCUSED).put(KeyStroke.getKeyStroke(key, modifiers), name);
            table.getActionMap().put(name, new AbstractAction() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    action.run();
                }
            });
        }
    }

    /** In-cell editor that shows the raw input (formula text) instead of the computed value. */
    private static final class FormulaCellEditor extends DefaultCellEditor {
        private final SpreadsheetModel model;

        FormulaCellEditor(SpreadsheetModel model, Font font) {
            super(new JTextField());
            this.model = model;
            setClickCountToStart(2);
            JTextField field = (JTextField) getComponent();
            field.setFont(font);
            field.setBorder(BorderFactory.createLineBorder(GREEN, 2));
        }

        @Override
        public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected,
                                                     int row, int column) {
            return super.getTableCellEditorComponent(table, model.getInput(row, column), isSelected, row, column);
        }
    }

    // ==================================================================
    // Actions on the selection
    // ==================================================================

    private void applyToSelection(Consumer<CellStyle> change) {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        int[] rows = sv.table.getSelectedRows();
        int[] cols = sv.table.getSelectedColumns();
        sv.model.beginBatch();
        try {
            for (int r : rows) for (int c : cols) sv.model.updateStyle(r, c, change);
        } finally {
            sv.model.endBatch();
        }
        sv.table.requestFocusInWindow();
    }

    /** Excel behaviour: if every selected cell already has the style, remove it; otherwise apply it. */
    private void toggleStyle(boolean bold) {
        SheetView sv = currentSheet();
        if (sv == null) return;
        int[] rows = sv.table.getSelectedRows();
        int[] cols = sv.table.getSelectedColumns();
        if (rows.length == 0 || cols.length == 0) return;
        boolean all = true;
        outer:
        for (int r : rows) {
            for (int c : cols) {
                CellStyle st = sv.model.getStyle(r, c);
                boolean on = st != null && (bold ? st.bold : st.italic);
                if (!on) {
                    all = false;
                    break outer;
                }
            }
        }
        final boolean value = !all;
        applyToSelection(st -> {
            if (bold) st.bold = value;
            else st.italic = value;
        });
    }

    private void align(int alignment) {
        applyToSelection(st -> st.alignment = alignment);
    }

    private void pickFillColor() {
        if (!guardEditable(currentSheet())) return;
        Color c = javax.swing.JColorChooser.showDialog(this, "Choose Background Color", currentBackground);
        if (c == null) return;
        currentBackground = c;
        fillChip.setBackground(c);
        applyToSelection(st -> st.background = c);
    }

    private void clearSelection() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        int[] rows = sv.table.getSelectedRows();
        int[] cols = sv.table.getSelectedColumns();
        sv.model.beginBatch();
        try {
            for (int r : rows) for (int c : cols) sv.model.clearContents(r, c);
        } finally {
            sv.model.endBatch();
        }
    }

    /**
     * AutoSum. Single cell: sums the contiguous numbers above (or, failing that, to the left).
     * Multi-cell selection: puts a SUM for each selected column in the row below the selection.
     */
    private void autoSum() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        SpreadsheetModel m = sv.model;
        int[] rows = sv.table.getSelectedRows();
        int[] cols = sv.table.getSelectedColumns();
        if (rows.length == 0 || cols.length == 0) {
            status("Select a cell first");
            return;
        }
        int r1 = rows[0], r2 = rows[rows.length - 1], c1 = cols[0], c2 = cols[cols.length - 1];

        if (r1 == r2 && c1 == c2) {
            int top = r1;
            while (top - 1 >= 0 && m.isNumeric(top - 1, c1)) top--;
            if (top < r1) {
                m.setInput(r1, c1, "=SUM(" + FormulaEngine.cellRef(top, c1) + ":" + FormulaEngine.cellRef(r1 - 1, c1) + ")");
                return;
            }
            int left = c1;
            while (left - 1 >= 0 && m.isNumeric(r1, left - 1)) left--;
            if (left < c1) {
                m.setInput(r1, c1, "=SUM(" + FormulaEngine.cellRef(r1, left) + ":" + FormulaEngine.cellRef(r1, c1 - 1) + ")");
                return;
            }
            m.setInput(r1, c1, "=SUM()");
            status("No adjacent numbers found - type the range inside SUM( )");
            formulaField.requestFocusInWindow();
            formulaField.setCaretPosition(formulaField.getText().length() - 1);
        } else {
            int target = r2 + 1;
            if (target >= m.getRowCount()) {
                status("No room below the selection for the totals");
                return;
            }
            m.beginBatch();
            try {
                for (int c : cols) {
                    m.setInput(target, c, "=SUM(" + FormulaEngine.cellRef(r1, c) + ":" + FormulaEngine.cellRef(r2, c) + ")");
                }
            } finally {
                m.endBatch();
            }
        }
    }

    // ==================================================================
    // Dialogs
    // ==================================================================

    private void openInsertDialog() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        int[] rows = sv.table.getSelectedRows();
        int[] cols = sv.table.getSelectedColumns();
        int r = rows.length > 0 ? rows[0] : 0;
        int c = cols.length > 0 ? cols[0] : 0;
        new InsertDialog(this, sv.model, r, c, () -> {
            syncFormulaBar();
            sv.table.requestFocusInWindow();
            status("Inserted at " + FormulaEngine.cellRef(r, c));
        }).setVisible(true);
    }

    private void openMacroIde() {
        if (macroDialog == null) {
            macroDialog = new MacroIDEDialog(this, () -> {
                SheetView sv = currentSheet();
                return sv == null ? null : sv.model;
            });
        }
        macroDialog.setVisible(true);
        macroDialog.toFront();
    }

    // ==================================================================
    // CSV import / export
    // ==================================================================

    // ==================================================================
    // Sheet protection (Review menu)
    // ==================================================================

    /** @return false (with a hint in the status bar) when the sheet is protected. */
    private boolean guardEditable(SheetView sv) {
        if (sv != null && sv.model.isProtected()) {
            status("This sheet is protected. Use Review > Unprotect Sheet to make changes.");
            Toolkit.getDefaultToolkit().beep();
            return false;
        }
        return true;
    }

    private void updateProtectUi(SheetView sv) {
        boolean prot = sv != null && sv.model.isProtected();
        if (protectButton != null) protectButton.setText(prot ? "Unprotect Sheet" : "Protect Sheet");
        protLabel.setText(prot ? "SHEET PROTECTED" : "");
        formulaField.setEditable(!prot);
    }

    private void refreshTabLocks() {
        for (int i = 0; i < tabs.getTabCount(); i++) {
            Component comp = tabs.getComponentAt(i);
            Component header = tabs.getTabComponentAt(i);
            if (comp instanceof SheetView && header instanceof TabHeader) {
                ((TabHeader) header).setLocked(((SheetView) comp).model.isProtected());
            }
        }
    }

    private void toggleProtection() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        if (sv.model.isProtected()) unprotectSheetDialog();
        else protectSheetDialog();
    }

    private void protectSheetDialog() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        if (sv.model.isProtected()) {
            JOptionPane.showMessageDialog(this, "This sheet is already protected.", "Protect Sheet",
                    JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        char[] pw = askNewPassword("Protect Sheet",
                "Password (optional - leave blank to protect without a password):",
                "Protected sheets cannot be edited, formatted, cleared, changed by macros, or have pictures moved.");
        if (pw == null) return;
        if (sv.table.isEditing()) sv.table.getCellEditor().cancelCellEditing();
        sv.model.getProtection().protect(pw);
        Arrays.fill(pw, '\0');
        afterProtectionChange(sv, "Sheet protected");
    }

    private void unprotectSheetDialog() {
        SheetView sv = currentSheet();
        if (sv == null || !sv.model.isProtected()) return;
        char[] pw = null;
        if (sv.model.getProtection().hasPassword()) {
            pw = askPassword("Unprotect Sheet", "Password for \"" + tabs.getTitleAt(tabs.getSelectedIndex()) + "\":");
            if (pw == null) return;
        }
        boolean ok = sv.model.getProtection().unprotect(pw);
        if (pw != null) Arrays.fill(pw, '\0');
        if (!ok) {
            JOptionPane.showMessageDialog(this, "Incorrect password.", "Unprotect Sheet", JOptionPane.ERROR_MESSAGE);
            return;
        }
        afterProtectionChange(sv, "Sheet unprotected");
    }

    private void afterProtectionChange(SheetView sv, String message) {
        refreshTabLocks();
        markDirty();
        sv.table.repaint();
        syncFormulaBar();
        status(message);
        sv.table.requestFocusInWindow();
    }

    // ---- password prompts ----

    /** Asks for a password once. @return the characters, or null if cancelled. */
    private char[] askPassword(String title, String message) {
        JPasswordField field = new JPasswordField(18);
        focusOnShow(field);
        JPanel panel = new JPanel(new GridLayout(0, 1, 4, 4));
        panel.add(new JLabel(message));
        panel.add(field);
        int answer = JOptionPane.showConfirmDialog(this, panel, title,
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        return answer == JOptionPane.OK_OPTION ? field.getPassword() : null;
    }

    /**
     * Asks for a new password with confirmation.
     * @return the characters (possibly empty), or null if cancelled or the two entries differ
     */
    private char[] askNewPassword(String title, String message, String note) {
        JPasswordField first = new JPasswordField(18);
        JPasswordField second = new JPasswordField(18);
        focusOnShow(first);
        JPanel panel = new JPanel(new GridLayout(0, 1, 4, 4));
        panel.add(new JLabel(message));
        panel.add(first);
        panel.add(new JLabel("Confirm password:"));
        panel.add(second);
        panel.add(new JLabel("<html><small>" + note + "</small></html>"));
        int answer = JOptionPane.showConfirmDialog(this, panel, title,
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (answer != JOptionPane.OK_OPTION) return null;
        char[] a = first.getPassword();
        char[] b = second.getPassword();
        boolean same = Arrays.equals(a, b);
        Arrays.fill(b, '\0');
        if (!same) {
            Arrays.fill(a, '\0');
            JOptionPane.showMessageDialog(this, "The passwords do not match.", title, JOptionPane.ERROR_MESSAGE);
            return null;
        }
        return a;
    }

    /** Puts the keyboard cursor in the field when its dialog appears. */
    private static void focusOnShow(JComponent field) {
        field.addAncestorListener(new AncestorListener() {
            @Override
            public void ancestorAdded(AncestorEvent event) {
                SwingUtilities.invokeLater(field::requestFocusInWindow);
            }

            @Override
            public void ancestorRemoved(AncestorEvent event) {
            }

            @Override
            public void ancestorMoved(AncestorEvent event) {
            }
        });
    }

    /** Small padlock drawn on the tab of a protected sheet. */
    private static final class LockIcon implements Icon {
        @Override
        public void paintIcon(Component c, Graphics g, int x, int y) {
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.translate(x, y);
            g2.setColor(new Color(0x8A6D00));
            g2.setStroke(new BasicStroke(1.6f));
            g2.drawArc(2, 0, 7, 8, 0, 180);
            g2.drawLine(2, 4, 2, 6);
            g2.drawLine(9, 4, 9, 6);
            g2.setColor(new Color(0xD9A900));
            g2.fillRoundRect(0, 6, 11, 7, 3, 3);
            g2.dispose();
        }

        @Override
        public int getIconWidth() {
            return 11;
        }

        @Override
        public int getIconHeight() {
            return 13;
        }
    }

    // ==================================================================
    // Pictures
    // ==================================================================

    private void insertPictureFromFile() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter(new FileNameExtensionFilter("Pictures (png, jpg, gif, bmp)",
                "png", "jpg", "jpeg", "gif", "bmp"));
        if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;
        try {
            BufferedImage img = ImageIO.read(fc.getSelectedFile());
            if (img == null) {
                JOptionPane.showMessageDialog(this, "That file is not a supported picture (use PNG, JPG, GIF or BMP).",
                        "Insert Picture", JOptionPane.ERROR_MESSAGE);
                return;
            }
            placePicture(sv, img);
        } catch (IOException | RuntimeException ex) {
            JOptionPane.showMessageDialog(this, "Could not read the picture:\n" + ex.getMessage(),
                    "Insert Picture", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void insertPictureFromClipboard() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        try {
            Clipboard cb = Toolkit.getDefaultToolkit().getSystemClipboard();
            if (!cb.isDataFlavorAvailable(DataFlavor.imageFlavor)) {
                JOptionPane.showMessageDialog(this, "There is no picture on the clipboard.",
                        "Insert Picture", JOptionPane.INFORMATION_MESSAGE);
                return;
            }
            java.awt.Image im = (java.awt.Image) cb.getData(DataFlavor.imageFlavor);
            javax.swing.ImageIcon icon = new javax.swing.ImageIcon(im); // makes sure it is fully loaded
            int w = icon.getIconWidth();
            int h = icon.getIconHeight();
            if (w <= 0 || h <= 0) throw new IOException("The clipboard picture is empty");
            BufferedImage bi = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g = bi.createGraphics();
            g.drawImage(im, 0, 0, null);
            g.dispose();
            placePicture(sv, bi);
        } catch (java.awt.datatransfer.UnsupportedFlavorException | IOException | RuntimeException ex) {
            JOptionPane.showMessageDialog(this, "Could not paste the picture:\n" + ex.getMessage(),
                    "Insert Picture", JOptionPane.ERROR_MESSAGE);
        }
    }

    /** Puts the picture at the active cell and selects it so it can be moved or resized. */
    private void placePicture(SheetView sv, BufferedImage img) throws IOException {
        int r = Math.max(0, sv.table.getSelectionModel().getLeadSelectionIndex());
        int c = Math.max(0, sv.table.getColumnModel().getSelectionModel().getLeadSelectionIndex());
        Rectangle cell = sv.table.getCellRect(r, c, true);
        SheetImage si = SheetImage.create(img, cell.x, cell.y);
        sv.model.addImage(si);
        sv.layer.select(si);
        markDirty();
        sv.table.requestFocusInWindow();
        status("Picture inserted at " + FormulaEngine.cellRef(r, c)
                + " - drag to move, drag the corner square to resize, Delete to remove");
    }

    private void deleteSelectedPicture() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        if (sv.layer.deleteSelected()) status("Picture deleted");
        else status("Click a picture to select it first");
    }

    // ==================================================================
    // Workbook files (.spx), open password, unsaved-changes tracking
    // ==================================================================

    private void markDirty() {
        if (loading || dirty) return;
        dirty = true;
        updateTitle();
    }

    private void updateTitle() {
        String name = currentFile == null ? "Untitled" : currentFile.getName();
        setTitle("Spreadsheet Pro - " + name + (dirty ? " *" : ""));
    }

    /** @return true if it is fine to replace the current workbook (saved, discarded or unchanged). */
    private boolean confirmDiscard() {
        if (!dirty) return true;
        String name = currentFile == null ? "Untitled" : currentFile.getName();
        int answer = JOptionPane.showConfirmDialog(this, "Save changes to " + name + "?", "Unsaved Changes",
                JOptionPane.YES_NO_CANCEL_OPTION, JOptionPane.WARNING_MESSAGE);
        if (answer == JOptionPane.YES_OPTION) return saveWorkbook();
        return answer == JOptionPane.NO_OPTION;
    }

    private void exitApp() {
        if (!confirmDiscard()) return;
        dispose();
        System.exit(0);
    }

    private void newWorkbook() {
        if (!confirmDiscard()) return;
        loading = true;
        try {
            tabs.removeAll();
            sheetCounter = 0;
            addSheet(nextSheetName());
        } finally {
            loading = false;
        }
        currentFile = null;
        if (workbookPassword != null) Arrays.fill(workbookPassword, '\0');
        workbookPassword = null;
        dirty = false;
        updateTitle();
        refreshTabLocks();
        syncFormulaBar();
        status("New workbook");
    }

    private JFileChooser workbookChooser() {
        JFileChooser fc = new JFileChooser(currentFile != null ? currentFile.getParentFile() : null);
        fc.setFileFilter(new FileNameExtensionFilter("Spreadsheet Pro workbook (*.spx)", WorkbookIO.EXTENSION));
        return fc;
    }

    private List<WorkbookIO.SheetData> collectSheets() {
        List<WorkbookIO.SheetData> list = new ArrayList<>();
        for (int i = 0; i < tabs.getTabCount(); i++) {
            Component comp = tabs.getComponentAt(i);
            if (comp instanceof SheetView) {
                list.add(new WorkbookIO.SheetData(tabs.getTitleAt(i), ((SheetView) comp).model));
            }
        }
        return list;
    }

    /** Ctrl+S. @return true if the workbook was saved. */
    private boolean saveWorkbook() {
        if (currentFile == null) return saveWorkbookAs();
        return writeWorkbook(currentFile);
    }

    private boolean saveWorkbookAs() {
        JFileChooser fc = workbookChooser();
        fc.setSelectedFile(new File(currentFile != null ? currentFile.getName() : "Workbook.spx"));
        if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return false;
        File f = fc.getSelectedFile();
        if (!f.getName().toLowerCase().endsWith("." + WorkbookIO.EXTENSION)) {
            f = new File(f.getParentFile(), f.getName() + "." + WorkbookIO.EXTENSION);
        }
        if (f.exists() && JOptionPane.showConfirmDialog(this, "\"" + f.getName() + "\" already exists. Replace it?",
                "Save As", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE) != JOptionPane.YES_OPTION) {
            return false;
        }
        if (!writeWorkbook(f)) return false;
        currentFile = f;
        updateTitle();
        return true;
    }

    private boolean writeWorkbook(File f) {
        SheetView sv = currentSheet();
        if (sv != null && sv.table.isEditing()) sv.table.getCellEditor().stopCellEditing();
        try {
            WorkbookIO.save(f, collectSheets(), tabs.getSelectedIndex(), workbookPassword);
            dirty = false;
            updateTitle();
            status("Saved " + f.getName() + (workbookPassword != null ? " (encrypted)" : ""));
            return true;
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Could not save the workbook:\n" + ex.getMessage(),
                    "Save", JOptionPane.ERROR_MESSAGE);
            return false;
        }
    }

    private void openWorkbook() {
        if (!confirmDiscard()) return;
        JFileChooser fc = workbookChooser();
        if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;
        File f = fc.getSelectedFile();
        char[] pw = null;
        WorkbookIO.WorkbookData data;
        while (true) {
            try {
                data = WorkbookIO.load(f, pw);
                break;
            } catch (WorkbookIO.PasswordRequiredException | WorkbookIO.BadPasswordException ex) {
                boolean wrong = ex instanceof WorkbookIO.BadPasswordException;
                pw = askPassword("Open Workbook",
                        wrong ? "Incorrect password. Try again:" : "\"" + f.getName() + "\" needs a password to open:");
                if (pw == null) return;
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Could not open the workbook:\n" + ex.getMessage(),
                        "Open", JOptionPane.ERROR_MESSAGE);
                return;
            }
        }

        loading = true;
        try {
            tabs.removeAll();
            for (WorkbookIO.SheetData sd : data.sheets) addSheet(sd.name, sd.model);
            tabs.setSelectedIndex(data.activeIndex);
        } finally {
            loading = false;
        }
        sheetCounter = tabs.getTabCount();
        currentFile = f;
        if (workbookPassword != null) Arrays.fill(workbookPassword, '\0');
        workbookPassword = pw;
        dirty = false;
        updateTitle();
        refreshTabLocks();
        syncFormulaBar();
        status("Opened " + f.getName());
    }

    private void setOpenPassword() {
        char[] pw = askNewPassword("Set Open Password",
                "Password needed to open this workbook (leave blank to remove it):",
                "The workbook is encrypted when you save it. If you forget the password the file cannot be recovered.");
        if (pw == null) return;
        if (workbookPassword != null) Arrays.fill(workbookPassword, '\0');
        workbookPassword = pw.length == 0 ? null : pw;
        markDirty();
        status(pw.length == 0 ? "Open password removed - save to apply"
                : "Open password set - save the workbook to encrypt it");
    }

    private void exportCsv() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        JFileChooser fc = new JFileChooser();
        fc.setSelectedFile(new File(tabs.getTitleAt(tabs.getSelectedIndex()) + ".csv"));
        if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return;
        File file = fc.getSelectedFile();
        try {
            Files.write(file.toPath(), toCsv(sv.model).getBytes(StandardCharsets.UTF_8));
            status("Exported " + file.getName());
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Could not write file:\n" + ex.getMessage(),
                    "Export CSV", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void importCsv() {
        JFileChooser fc = new JFileChooser();
        if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;
        File file = fc.getSelectedFile();
        try {
            String content = new String(Files.readAllBytes(file.toPath()), StandardCharsets.UTF_8);
            String name = file.getName().replaceFirst("\\.[^.]*$", "");
            SheetView sv = addSheet(name.isEmpty() ? nextSheetName() : name);
            List<List<String>> rows = parseCsv(content);
            sv.model.beginBatch();
            try {
                for (int r = 0; r < rows.size(); r++) {
                    List<String> row = rows.get(r);
                    for (int c = 0; c < row.size(); c++) sv.model.setInput(r, c, row.get(c));
                }
            } finally {
                sv.model.endBatch();
            }
            status("Imported " + file.getName());
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Could not read file:\n" + ex.getMessage(),
                    "Import CSV", JOptionPane.ERROR_MESSAGE);
        }
    }

    private static String toCsv(SpreadsheetModel m) {
        StringBuilder sb = new StringBuilder();
        int rows = m.getUsedRowCount();
        int cols = m.getUsedColumnCount();
        for (int r = 0; r < rows; r++) {
            for (int c = 0; c < cols; c++) {
                if (c > 0) sb.append(',');
                String v = m.getInput(r, c);
                if (v.contains(",") || v.contains("\"") || v.contains("\n")) {
                    v = "\"" + v.replace("\"", "\"\"") + "\"";
                }
                sb.append(v);
            }
            sb.append("\r\n");
        }
        return sb.toString();
    }

    static List<List<String>> parseCsv(String s) {
        List<List<String>> rows = new ArrayList<>();
        List<String> row = new ArrayList<>();
        StringBuilder cell = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (inQuotes) {
                if (ch == '"') {
                    if (i + 1 < s.length() && s.charAt(i + 1) == '"') {
                        cell.append('"');
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    cell.append(ch);
                }
            } else if (ch == '"') {
                inQuotes = true;
            } else if (ch == ',') {
                row.add(cell.toString());
                cell.setLength(0);
            } else if (ch == '\n') {
                row.add(cell.toString());
                cell.setLength(0);
                rows.add(row);
                row = new ArrayList<>();
            } else if (ch != '\r') {
                cell.append(ch);
            }
        }
        if (cell.length() > 0 || !row.isEmpty()) {
            row.add(cell.toString());
            rows.add(row);
        }
        return rows;
    }
}
