import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.util.function.Supplier;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JTextArea;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;

/**
 * Lightweight "VBA" editor: a code editor with line numbers, a Run button
 * and an output console. Scripts run through {@link MacroEngine} against the
 * worksheet that is active when Run is pressed.
 */
public class MacroIDEDialog extends JDialog {

    private static final Color GREEN = new Color(0x107C41);

    private static final String SAMPLE =
            "Sub FillSquares()\n"
            + "    ' Header row\n"
            + "    Range(\"A1\").Value = \"N\"\n"
            + "    Range(\"B1\").Value = \"Square\"\n"
            + "    Range(\"A1:B1\").Font.Bold = True\n"
            + "    Range(\"A1:B1\").Interior.Color = RGB(198, 229, 209)\n"
            + "\n"
            + "    ' Loop through cells and set values programmatically\n"
            + "    For i = 1 To 10\n"
            + "        Cells(i + 1, 1).Value = i\n"
            + "        Cells(i + 1, 2).Value = i * i\n"
            + "        If i Mod 2 = 0 Then\n"
            + "            Cells(i + 1, 1).Interior.Color = RGB(233, 245, 238)\n"
            + "        End If\n"
            + "    Next i\n"
            + "\n"
            + "    Range(\"A12\").Value = \"Total\"\n"
            + "    Range(\"B12\").Formula = \"=SUM(B2:B11)\"\n"
            + "    Range(\"A12:B12\").Font.Bold = True\n"
            + "    MsgBox \"Done! Sum of squares is in B12\"\n"
            + "End Sub\n";

    private final Supplier<SpreadsheetModel> modelSupplier;
    private final JTextArea editor = new JTextArea(SAMPLE);
    private final JTextArea lineNumbers = new JTextArea("1");
    private final JTextArea console = new JTextArea();

    public MacroIDEDialog(JFrame owner, Supplier<SpreadsheetModel> modelSupplier) {
        super(owner, "VBA Macro IDE", false);
        this.modelSupplier = modelSupplier;

        Font mono = new Font(Font.MONOSPACED, Font.PLAIN, 14);
        editor.setFont(mono);
        editor.setTabSize(4);
        editor.setMargin(new java.awt.Insets(4, 6, 4, 6));
        lineNumbers.setFont(mono);
        lineNumbers.setEditable(false);
        lineNumbers.setFocusable(false);
        lineNumbers.setBackground(new Color(0xF0F0F0));
        lineNumbers.setForeground(new Color(0x888888));
        lineNumbers.setMargin(new java.awt.Insets(4, 6, 4, 6));
        editor.getDocument().addDocumentListener(new DocumentListener() {
            public void insertUpdate(DocumentEvent e) { updateLineNumbers(); }
            public void removeUpdate(DocumentEvent e) { updateLineNumbers(); }
            public void changedUpdate(DocumentEvent e) { updateLineNumbers(); }
        });
        updateLineNumbers();

        JScrollPane editorScroll = new JScrollPane(editor);
        editorScroll.setRowHeaderView(lineNumbers);

        console.setEditable(false);
        console.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 13));
        console.setBackground(new Color(0x1E1E1E));
        console.setForeground(new Color(0xD4D4D4));
        console.setMargin(new java.awt.Insets(4, 6, 4, 6));
        JScrollPane consoleScroll = new JScrollPane(console);
        consoleScroll.setPreferredSize(new Dimension(100, 150));

        JSplitPane split = new JSplitPane(JSplitPane.VERTICAL_SPLIT, editorScroll, consoleScroll);
        split.setResizeWeight(0.75);

        JButton run = new JButton("Run Macro");
        run.setBackground(GREEN);
        run.setForeground(Color.WHITE);
        run.addActionListener(e -> runMacro());
        JButton sample = new JButton("Load Sample");
        sample.addActionListener(e -> editor.setText(SAMPLE));
        JButton clear = new JButton("Clear Output");
        clear.addActionListener(e -> console.setText(""));
        JButton close = new JButton("Close");
        close.addActionListener(e -> setVisible(false));

        JPanel bar = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 6));
        bar.add(run);
        bar.add(sample);
        bar.add(clear);
        bar.add(close);
        JLabel hint = new JLabel("  Cells(r,c).Value, Range(\"A1:B3\").Formula, For/Next, If/Else, MsgBox, RGB()");
        hint.setForeground(Color.GRAY);
        bar.add(hint);
        bar.setBorder(BorderFactory.createMatteBorder(0, 0, 1, 0, new Color(0xD0D0D0)));

        add(bar, BorderLayout.NORTH);
        add(split, BorderLayout.CENTER);
        setSize(780, 580);
        setLocationRelativeTo(owner);
    }

    private void updateLineNumbers() {
        int lines = Math.max(1, editor.getLineCount());
        StringBuilder sb = new StringBuilder();
        for (int i = 1; i <= lines; i++) {
            if (i > 1) sb.append('\n');
            sb.append(i);
        }
        lineNumbers.setText(sb.toString());
    }

    private void print(String s) {
        console.append(s + "\n");
        console.setCaretPosition(console.getDocument().getLength());
    }

    private void runMacro() {
        SpreadsheetModel model = modelSupplier.get();
        if (model == null) {
            print("! No active worksheet.");
            return;
        }
        print("> Running macro...");
        MacroEngine engine = new MacroEngine(model, this::print);
        try {
            int steps = engine.run(editor.getText());
            print("> Finished OK (" + steps + " steps)");
        } catch (MacroEngine.MacroException ex) {
            print("! " + ex.getMessage());
        } catch (RuntimeException ex) {
            print("! Internal error: " + ex);
        }
    }
}
