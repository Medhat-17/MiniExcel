import java.awt.Color;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Random;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.SwingConstants;

/**
 * A small interpreter for a useful subset of VBA, used by {@link MacroIDEDialog}.
 *
 * Supported statements:
 * <pre>
 *   Sub Name() / End Sub          Dim x As Integer          ' comments
 *   x = expression                For i = 1 To 10 [Step 2] ... Next [i]
 *   If cond Then ... [ElseIf cond Then ...] [Else ...] End If   (also single-line If)
 *   Exit For                      MsgBox expr                Debug.Print expr
 *   Cells(row, col).Value = expr  Range("A1").Value = expr   Range("A1:B3").Formula = "=SUM(A1:A5)"
 *   Range("A1").Font.Bold = True  .Font.Italic  .Font.Size  .Font.Color  .Interior.Color = RGB(r,g,b)
 *   Range("A1:C3").ClearContents  Range("A1:C3").Clear
 * </pre>
 * Expressions: + - * / \ Mod ^ &amp; comparisons, And Or Not, string / number literals,
 * Cells(r,c).Value reads, and the functions Int Abs Sqr Rnd Round Len UCase LCase
 * Str Val Mid RGB.
 */
public class MacroEngine {

    /** Error raised for any syntax or runtime problem in a macro. */
    public static class MacroException extends Exception {
        public MacroException(String message, int line) {
            super(line > 0 ? "Line " + line + ": " + message : message);
        }
    }

    private static final int MAX_STEPS = 500_000;
    private static final int CI = Pattern.CASE_INSENSITIVE;
    private static final Pattern IF_PATTERN = Pattern.compile("^if\\s+(.+?)\\s+then\\s*(.*)$", CI);
    private static final Pattern ELSEIF_PATTERN = Pattern.compile("^elseif\\s+(.+?)\\s+then\\s*$", CI);
    private static final Pattern END_IF = Pattern.compile("^end\\s+if\\b", CI);
    private static final Pattern END_SUB = Pattern.compile("^end\\s+(sub|function)\\b", CI);
    private static final Pattern SUB_START = Pattern.compile("^(public\\s+|private\\s+)?(sub|function)\\b.*", CI);
    private static final Pattern FOR_PATTERN =
            Pattern.compile("^for\\s+([a-z_]\\w*)\\s*=\\s*(.+?)\\s+to\\s+(.+?)(?:\\s+step\\s+(.+))?$", CI);
    private static final Pattern EXIT_FOR = Pattern.compile("^exit\\s+for\\b", CI);
    private static final Pattern A1 = Pattern.compile("^\\$?([A-Za-z]{1,3})\\$?(\\d+)$");
    private static final Pattern LEADING_NUMBER = Pattern.compile("^\\s*[-+]?(\\d+\\.?\\d*|\\.\\d+)");

    private static final Map<String, Double> CONSTANTS = new HashMap<>();

    static {
        CONSTANTS.put("xlleft", -4131.0);
        CONSTANTS.put("xlcenter", -4108.0);
        CONSTANTS.put("xlright", -4152.0);
        CONSTANTS.put("vbblack", 0.0);
        CONSTANTS.put("vbred", 255.0);
        CONSTANTS.put("vbgreen", 65280.0);
        CONSTANTS.put("vbblue", 16711680.0);
        CONSTANTS.put("vbyellow", 65535.0);
        CONSTANTS.put("vbwhite", 16777215.0);
    }

    private final SpreadsheetModel model;
    private final Consumer<String> out;
    private final Map<String, Object> vars = new HashMap<>();
    private final Random random = new Random();
    private String[] src = new String[0];
    private int currentLine = 0;

    private static final class LoopFrame {
        String var;
        double end;
        double step;
        int forLine;
    }

    private static final class Tok {
        final char type; // n=number s=string i=identifier o=operator
        final String text;

        Tok(char type, String text) {
            this.type = type;
            this.text = text;
        }
    }

    public MacroEngine(SpreadsheetModel model, Consumer<String> out) {
        this.model = model;
        this.out = out;
    }

    private MacroException err(String message) {
        return new MacroException(message, currentLine);
    }

    // ------------------------------------------------------------------
    // Program execution
    // ------------------------------------------------------------------

    /** Runs the macro; returns the number of executed steps. */
    public int run(String source) throws MacroException {
        if (model.isProtected()) {
            throw new MacroException("This sheet is protected. Unprotect it before running a macro.", 0);
        }
        vars.clear();
        String[] raw = source.split("\\r?\\n", -1);
        src = new String[raw.length];
        for (int i = 0; i < raw.length; i++) src[i] = stripComment(raw[i]).trim();

        validateBlocks();

        Deque<LoopFrame> loops = new ArrayDeque<>();
        int pc = 0;
        int steps = 0;
        model.beginBatch();
        try {
            while (pc < src.length) {
                currentLine = pc + 1;
                if (++steps > MAX_STEPS) throw err("Execution limit exceeded (infinite loop?)");
                String line = src[pc];
                if (line.isEmpty()) {
                    pc++;
                    continue;
                }
                String lower = line.toLowerCase(Locale.ROOT);

                if (END_SUB.matcher(line).find()) break;
                if (SUB_START.matcher(line).matches() || kw(lower, "option")) {
                    pc++;
                    continue;
                }
                if (kw(lower, "dim")) {
                    declare(line.substring(3));
                    pc++;
                    continue;
                }
                if (kw(lower, "for")) {
                    pc = execFor(line, pc, loops);
                    continue;
                }
                if (kw(lower, "next")) {
                    if (loops.isEmpty()) throw err("Next without For");
                    LoopFrame f = loops.peek();
                    double v = toNum(vars.get(f.var)) + f.step;
                    vars.put(f.var, v);
                    boolean again = f.step > 0 ? v <= f.end : v >= f.end;
                    if (again) {
                        pc = f.forLine + 1;
                    } else {
                        loops.pop();
                        pc++;
                    }
                    continue;
                }
                if (EXIT_FOR.matcher(line).find()) {
                    if (loops.isEmpty()) throw err("Exit For outside a For loop");
                    LoopFrame f = loops.pop();
                    pc = findMatchingNext(f.forLine) + 1;
                    continue;
                }
                if (kw(lower, "if")) {
                    Matcher m = IF_PATTERN.matcher(line);
                    if (!m.matches()) throw err("Syntax error in If statement");
                    boolean cond = toBool(eval(m.group(1)));
                    String rest = m.group(2).trim();
                    if (!rest.isEmpty()) { // single-line If
                        if (cond && EXIT_FOR.matcher(rest).find()) { // "If x Then Exit For"
                            if (loops.isEmpty()) throw err("Exit For outside a For loop");
                            pc = findMatchingNext(loops.pop().forLine) + 1;
                            continue;
                        }
                        if (cond) execStatement(rest);
                        pc++;
                    } else if (cond) {
                        pc++;
                    } else {
                        pc = skipToBranch(pc);
                    }
                    continue;
                }
                if (kw(lower, "elseif") || kw(lower, "else")) { // finished the true branch
                    pc = findEndIf(pc) + 1;
                    continue;
                }
                if (END_IF.matcher(line).find()) {
                    pc++;
                    continue;
                }
                execStatement(line);
                pc++;
            }
        } finally {
            model.endBatch();
        }
        return steps;
    }

    /** Compile-time style check: every For has a Next and every block If has an End If. */
    private void validateBlocks() throws MacroException {
        Deque<Integer> forStack = new ArrayDeque<>();
        Deque<Integer> ifStack = new ArrayDeque<>();
        for (int i = 0; i < src.length; i++) {
            String l = src[i];
            if (l.isEmpty()) continue;
            currentLine = i + 1;
            String lo = l.toLowerCase(Locale.ROOT);
            if (END_SUB.matcher(l).find()) break;
            if (kw(lo, "for")) forStack.push(i);
            else if (kw(lo, "next")) {
                if (forStack.isEmpty()) throw err("Next without For");
                forStack.pop();
            } else if (isBlockIf(l)) ifStack.push(i);
            else if (END_IF.matcher(l).find()) {
                if (ifStack.isEmpty()) throw err("End If without If");
                ifStack.pop();
            }
        }
        if (!forStack.isEmpty()) {
            currentLine = forStack.peek() + 1;
            throw err("For without Next");
        }
        if (!ifStack.isEmpty()) {
            currentLine = ifStack.peek() + 1;
            throw err("If without End If");
        }
    }

    private int execFor(String line, int pc, Deque<LoopFrame> loops) throws MacroException {
        Matcher m = FOR_PATTERN.matcher(line);
        if (!m.matches()) throw err("Syntax error in For statement");
        String var = m.group(1).toLowerCase(Locale.ROOT);
        double start = toNum(eval(m.group(2)));
        double end = toNum(eval(m.group(3)));
        double step = m.group(4) == null ? 1 : toNum(eval(m.group(4)));
        if (step == 0) throw err("Step cannot be zero");
        vars.put(var, start);
        if (step > 0 ? start > end : start < end) return findMatchingNext(pc) + 1; // zero iterations
        LoopFrame f = new LoopFrame();
        f.var = var;
        f.end = end;
        f.step = step;
        f.forLine = pc;
        loops.push(f);
        return pc + 1;
    }

    private void declare(String decl) {
        for (String part : decl.split(",")) {
            String p = part.trim();
            if (p.isEmpty()) continue;
            String name = p.split("\\s+")[0].replaceAll("\\(.*", "").toLowerCase(Locale.ROOT);
            boolean isString = p.toLowerCase(Locale.ROOT).contains(" as string");
            vars.put(name, isString ? (Object) "" : (Object) 0.0);
        }
    }

    // ------------------------------------------------------------------
    // Block scanning helpers
    // ------------------------------------------------------------------

    private static boolean kw(String lower, String word) {
        if (!lower.startsWith(word)) return false;
        if (lower.length() == word.length()) return true;
        char c = lower.charAt(word.length());
        return !(Character.isLetterOrDigit(c) || c == '_');
    }

    private boolean isBlockIf(String line) {
        Matcher m = IF_PATTERN.matcher(line);
        return m.matches() && m.group(2).trim().isEmpty();
    }

    /** From a false If/ElseIf at {@code pc}, returns the pc of the first statement to run. */
    private int skipToBranch(int pc) throws MacroException {
        int idx = pc;
        while (true) {
            idx = findNextBranch(idx);
            String l = src[idx];
            String lo = l.toLowerCase(Locale.ROOT);
            if (kw(lo, "elseif")) {
                currentLine = idx + 1;
                Matcher m = ELSEIF_PATTERN.matcher(l);
                if (!m.matches()) throw err("Syntax error in ElseIf statement");
                if (toBool(eval(m.group(1)))) return idx + 1;
            } else {
                return idx + 1; // Else or End If
            }
        }
    }

    private int findNextBranch(int start) throws MacroException {
        int depth = 0;
        for (int i = start + 1; i < src.length; i++) {
            String l = src[i];
            if (l.isEmpty()) continue;
            String lo = l.toLowerCase(Locale.ROOT);
            if (isBlockIf(l)) depth++;
            else if (END_IF.matcher(l).find()) {
                if (depth == 0) return i;
                depth--;
            } else if (depth == 0 && (kw(lo, "elseif") || kw(lo, "else"))) return i;
        }
        throw err("If without End If");
    }

    private int findEndIf(int start) throws MacroException {
        int depth = 0;
        for (int i = start + 1; i < src.length; i++) {
            String l = src[i];
            if (l.isEmpty()) continue;
            if (isBlockIf(l)) depth++;
            else if (END_IF.matcher(l).find()) {
                if (depth == 0) return i;
                depth--;
            }
        }
        throw err("If without End If");
    }

    private int findMatchingNext(int forLine) throws MacroException {
        int depth = 0;
        for (int i = forLine + 1; i < src.length; i++) {
            String lo = src[i].toLowerCase(Locale.ROOT);
            if (kw(lo, "for")) depth++;
            else if (kw(lo, "next")) {
                if (depth == 0) return i;
                depth--;
            }
        }
        throw err("For without Next");
    }

    private static String stripComment(String s) {
        boolean inString = false;
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if (c == '"') inString = !inString;
            else if (c == '\'' && !inString) return s.substring(0, i);
        }
        String t = s.trim().toLowerCase(Locale.ROOT);
        if (t.equals("rem") || t.startsWith("rem ")) return "";
        return s;
    }

    // ------------------------------------------------------------------
    // Simple statements
    // ------------------------------------------------------------------

    private Object eval(String expression) throws MacroException {
        Expr e = new Expr(tokenize(expression), 0);
        Object v = e.parseExpr();
        e.expectEnd();
        return v;
    }

    private void execStatement(String line) throws MacroException {
        List<Tok> toks = tokenize(line);
        if (toks.isEmpty()) return;
        Tok t0 = toks.get(0);
        if (t0.type == 'i') {
            String id = t0.text.toLowerCase(Locale.ROOT);
            if (id.equals("application")) return; // e.g. Application.ScreenUpdating = False
            if (id.equals("msgbox")) {
                out.accept("MsgBox: " + str(new Expr(toks, 1).parseExpr()));
                return;
            }
            if (id.equals("debug") && toks.size() > 3 && toks.get(2).text.equalsIgnoreCase("print")) {
                out.accept(str(new Expr(toks, 3).parseExpr()));
                return;
            }
            if (id.equals("cells") || id.equals("range")) {
                Expr e = new Expr(toks, 0);
                int[] ref = e.parseRef();
                List<String> props = new ArrayList<>();
                while (e.acceptOp(".")) props.add(e.expectIdent().toLowerCase(Locale.ROOT));
                if (e.acceptOp("=")) {
                    Object v = e.parseExpr();
                    e.expectEnd();
                    assign(ref, String.join(".", props), v);
                } else {
                    e.expectEnd();
                    invoke(ref, String.join(".", props));
                }
                return;
            }
            if (toks.size() >= 2 && toks.get(1).type == 'o' && toks.get(1).text.equals("=")) {
                Expr e = new Expr(toks, 2);
                Object v = e.parseExpr();
                e.expectEnd();
                vars.put(id, v);
                return;
            }
        }
        throw err("Syntax error: cannot understand '" + line + "'");
    }

    private void assign(int[] ref, String prop, Object v) throws MacroException {
        switch (prop) {
            case "":
            case "value":
            case "formula": {
                String text = inputString(v);
                for (int r = ref[0]; r <= ref[2]; r++)
                    for (int c = ref[1]; c <= ref[3]; c++) model.setInput(r, c, text);
                break;
            }
            case "font.bold": {
                final boolean b = toBool(v);
                forEachCell(ref, st -> st.bold = b);
                break;
            }
            case "font.italic": {
                final boolean b = toBool(v);
                forEachCell(ref, st -> st.italic = b);
                break;
            }
            case "font.size": {
                final int size = (int) toNum(v);
                forEachCell(ref, st -> st.fontSize = size);
                break;
            }
            case "font.color": {
                final Color col = fromVba(toNum(v));
                forEachCell(ref, st -> st.foreground = col);
                break;
            }
            case "interior.color": {
                final Color col = fromVba(toNum(v));
                forEachCell(ref, st -> st.background = col);
                break;
            }
            case "horizontalalignment": {
                int code = (int) toNum(v);
                final int align = code == -4108 ? SwingConstants.CENTER
                        : code == -4152 ? SwingConstants.RIGHT
                        : code == -4131 ? SwingConstants.LEFT : CellStyle.ALIGN_AUTO;
                forEachCell(ref, st -> st.alignment = align);
                break;
            }
            default:
                throw err("Unsupported property: " + prop);
        }
    }

    private void invoke(int[] ref, String method) throws MacroException {
        switch (method) {
            case "clearcontents":
                for (int r = ref[0]; r <= ref[2]; r++)
                    for (int c = ref[1]; c <= ref[3]; c++) model.clearContents(r, c);
                break;
            case "clear":
                for (int r = ref[0]; r <= ref[2]; r++)
                    for (int c = ref[1]; c <= ref[3]; c++) model.clearContents(r, c);
                forEachCell(ref, CellStyle::reset);
                break;
            default:
                throw err("Unsupported method: " + method);
        }
    }

    private void forEachCell(int[] ref, Consumer<CellStyle> change) {
        for (int r = ref[0]; r <= ref[2]; r++)
            for (int c = ref[1]; c <= ref[3]; c++) model.updateStyle(r, c, change);
    }

    private static Color fromVba(double value) {
        int n = (int) value;
        return new Color(n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF);
    }

    // ------------------------------------------------------------------
    // Value conversion
    // ------------------------------------------------------------------

    private double toNum(Object v) throws MacroException {
        if (v == null) return 0;
        if (v instanceof Double) return (Double) v;
        if (v instanceof Boolean) return ((Boolean) v) ? -1 : 0;
        String s = v.toString().trim();
        if (s.isEmpty()) return 0;
        try {
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            throw err("Type mismatch: '" + s + "' is not a number");
        }
    }

    private boolean toBool(Object v) throws MacroException {
        if (v instanceof Boolean) return (Boolean) v;
        if (v instanceof String) {
            String s = ((String) v).trim();
            if (s.equalsIgnoreCase("true")) return true;
            if (s.equalsIgnoreCase("false")) return false;
        }
        return toNum(v) != 0;
    }

    private static String str(Object v) {
        if (v == null) return "";
        if (v instanceof Double) return FormulaEngine.formatNumber((Double) v);
        if (v instanceof Boolean) return ((Boolean) v) ? "True" : "False";
        return v.toString();
    }

    private static String inputString(Object v) {
        if (v instanceof Boolean) return ((Boolean) v) ? "TRUE" : "FALSE";
        return str(v);
    }

    // ------------------------------------------------------------------
    // Tokenizer
    // ------------------------------------------------------------------

    private List<Tok> tokenize(String s) throws MacroException {
        List<Tok> toks = new ArrayList<>();
        int i = 0;
        int len = s.length();
        while (i < len) {
            char c = s.charAt(i);
            if (Character.isWhitespace(c)) {
                i++;
            } else if (Character.isDigit(c) || (c == '.' && i + 1 < len && Character.isDigit(s.charAt(i + 1)))) {
                int j = i;
                while (j < len && (Character.isDigit(s.charAt(j)) || s.charAt(j) == '.')) j++;
                toks.add(new Tok('n', s.substring(i, j)));
                i = j;
            } else if (c == '"') {
                StringBuilder sb = new StringBuilder();
                i++;
                while (true) {
                    if (i >= len) throw err("Unterminated string literal");
                    char d = s.charAt(i);
                    if (d == '"') {
                        if (i + 1 < len && s.charAt(i + 1) == '"') {
                            sb.append('"');
                            i += 2;
                            continue;
                        }
                        i++;
                        break;
                    }
                    sb.append(d);
                    i++;
                }
                toks.add(new Tok('s', sb.toString()));
            } else if (Character.isLetter(c) || c == '_') {
                int j = i;
                while (j < len && (Character.isLetterOrDigit(s.charAt(j)) || s.charAt(j) == '_')) j++;
                toks.add(new Tok('i', s.substring(i, j)));
                i = j;
            } else if (i + 1 < len && (s.startsWith("<>", i) || s.startsWith("<=", i) || s.startsWith(">=", i))) {
                toks.add(new Tok('o', s.substring(i, i + 2)));
                i += 2;
            } else if ("+-*/\\^&=<>(),.".indexOf(c) >= 0) {
                toks.add(new Tok('o', String.valueOf(c)));
                i++;
            } else {
                throw err("Unexpected character '" + c + "'");
            }
        }
        return toks;
    }

    // ------------------------------------------------------------------
    // Expression parser
    // ------------------------------------------------------------------

    private class Expr {
        private final List<Tok> t;
        private int p;

        Expr(List<Tok> t, int p) {
            this.t = t;
            this.p = p;
        }

        boolean atEnd() {
            return p >= t.size();
        }

        Tok peek() {
            return p < t.size() ? t.get(p) : null;
        }

        boolean peekOp(String op) {
            Tok k = peek();
            return k != null && k.type == 'o' && k.text.equals(op);
        }

        boolean acceptOp(String op) {
            if (peekOp(op)) {
                p++;
                return true;
            }
            return false;
        }

        boolean acceptWord(String w) {
            Tok k = peek();
            if (k != null && k.type == 'i' && k.text.equalsIgnoreCase(w)) {
                p++;
                return true;
            }
            return false;
        }

        void expectOp(String op) throws MacroException {
            if (!acceptOp(op)) throw err("Expected '" + op + "'");
        }

        String expectIdent() throws MacroException {
            Tok k = peek();
            if (k == null || k.type != 'i') throw err("Identifier expected");
            p++;
            return k.text;
        }

        void expectEnd() throws MacroException {
            if (!atEnd()) throw err("Unexpected '" + t.get(p).text + "'");
        }

        Object parseExpr() throws MacroException {
            Object left = andExpr();
            while (acceptWord("or")) {
                boolean r = toBool(andExpr());
                left = toBool(left) || r;
            }
            return left;
        }

        private Object andExpr() throws MacroException {
            Object left = notExpr();
            while (acceptWord("and")) {
                boolean r = toBool(notExpr());
                left = toBool(left) && r;
            }
            return left;
        }

        private Object notExpr() throws MacroException {
            if (acceptWord("not")) return !toBool(notExpr());
            return cmpExpr();
        }

        private Object cmpExpr() throws MacroException {
            Object left = concat();
            String[] ops = {"<>", "<=", ">=", "=", "<", ">"};
            for (String op : ops) {
                if (acceptOp(op)) return compare(left, concat(), op);
            }
            return left;
        }

        private boolean compare(Object a, Object b, String op) throws MacroException {
            int cmp;
            if (a instanceof String && b instanceof String) cmp = ((String) a).compareToIgnoreCase((String) b);
            else cmp = Double.compare(toNum(a), toNum(b));
            switch (op) {
                case "=": return cmp == 0;
                case "<>": return cmp != 0;
                case "<": return cmp < 0;
                case ">": return cmp > 0;
                case "<=": return cmp <= 0;
                default: return cmp >= 0;
            }
        }

        private Object concat() throws MacroException {
            Object left = addSub();
            while (acceptOp("&")) left = str(left) + str(addSub());
            return left;
        }

        private Object addSub() throws MacroException {
            Object left = mulDiv();
            while (true) {
                if (acceptOp("+")) {
                    Object r = mulDiv();
                    if (left instanceof String && r instanceof String) left = (String) left + r;
                    else left = toNum(left) + toNum(r);
                } else if (acceptOp("-")) {
                    left = toNum(left) - toNum(mulDiv());
                } else {
                    return left;
                }
            }
        }

        private Object mulDiv() throws MacroException {
            Object left = unary();
            while (true) {
                if (acceptOp("*")) {
                    left = toNum(left) * toNum(unary());
                } else if (acceptOp("/")) {
                    double d = toNum(unary());
                    if (d == 0) throw err("Division by zero");
                    left = toNum(left) / d;
                } else if (acceptOp("\\")) {
                    double d = toNum(unary());
                    if (Math.floor(d) == 0) throw err("Division by zero");
                    left = (double) ((long) toNum(left) / (long) d);
                } else if (acceptWord("mod")) {
                    long d = (long) toNum(unary());
                    if (d == 0) throw err("Division by zero");
                    left = (double) ((long) toNum(left) % d);
                } else {
                    return left;
                }
            }
        }

        private Object unary() throws MacroException {
            if (acceptOp("-")) return -toNum(unary());
            if (acceptOp("+")) return unary();
            return power();
        }

        private Object power() throws MacroException {
            Object base = primary();
            if (acceptOp("^")) return Math.pow(toNum(base), toNum(unary()));
            return base;
        }

        private Object primary() throws MacroException {
            Tok tok = peek();
            if (tok == null) throw err("Unexpected end of expression");
            if (tok.type == 'n') {
                p++;
                try {
                    return Double.parseDouble(tok.text);
                } catch (NumberFormatException e) {
                    throw err("Invalid number '" + tok.text + "'");
                }
            }
            if (tok.type == 's') {
                p++;
                return tok.text;
            }
            if (acceptOp("(")) {
                Object v = parseExpr();
                expectOp(")");
                return v;
            }
            if (tok.type != 'i') throw err("Unexpected '" + tok.text + "'");

            String name = tok.text.toLowerCase(Locale.ROOT);
            if (name.equals("true")) {
                p++;
                return Boolean.TRUE;
            }
            if (name.equals("false")) {
                p++;
                return Boolean.FALSE;
            }
            if (name.equals("cells") || name.equals("range")) {
                int[] ref = parseRef();
                String prop = "value";
                if (acceptOp(".")) prop = expectIdent().toLowerCase(Locale.ROOT);
                return readCell(ref[0], ref[1], prop);
            }
            p++;
            if (acceptOp("(")) {
                List<Object> args = new ArrayList<>();
                if (!acceptOp(")")) {
                    do {
                        args.add(parseExpr());
                    } while (acceptOp(","));
                    expectOp(")");
                }
                return callFunction(name, args);
            }
            if (name.equals("rnd")) return random.nextDouble();
            Double constant = CONSTANTS.get(name);
            if (constant != null) return constant;
            if (vars.containsKey(name)) return vars.get(name);
            throw err("Variable not defined: " + tok.text);
        }

        /** Parses Cells(r, c) or Range("A1[:B3]") and returns {row0, col0, row1, col1} (0-based). */
        int[] parseRef() throws MacroException {
            String name = expectIdent().toLowerCase(Locale.ROOT);
            expectOp("(");
            int r0, c0, r1, c1;
            if (name.equals("cells")) {
                Object ro = parseExpr();
                expectOp(",");
                Object co = parseExpr();
                expectOp(")");
                r0 = r1 = (int) Math.round(toNum(ro)) - 1;
                c0 = c1 = (co instanceof String && !((String) co).matches("\\d+"))
                        ? FormulaEngine.parseColumn((String) co)
                        : (int) Math.round(toNum(co)) - 1;
            } else {
                String a = str(parseExpr());
                String b = null;
                if (acceptOp(",")) b = str(parseExpr());
                expectOp(")");
                if (b == null && a.contains(":")) {
                    String[] parts = a.split(":");
                    if (parts.length != 2) throw err("Invalid range '" + a + "'");
                    a = parts[0];
                    b = parts[1];
                }
                int[] first = parseA1(a);
                int[] second = b == null ? first : parseA1(b);
                r0 = Math.min(first[0], second[0]);
                r1 = Math.max(first[0], second[0]);
                c0 = Math.min(first[1], second[1]);
                c1 = Math.max(first[1], second[1]);
            }
            if (!model.inBounds(r0, c0) || !model.inBounds(r1, c1)) {
                throw err("Subscript out of range: " + FormulaEngine.cellRef(Math.max(r0, 0), Math.max(c0, 0)));
            }
            return new int[]{r0, c0, r1, c1};
        }

        private int[] parseA1(String ref) throws MacroException {
            Matcher m = A1.matcher(ref.trim());
            if (!m.matches()) throw err("Invalid cell reference '" + ref + "'");
            return new int[]{Integer.parseInt(m.group(2)) - 1, FormulaEngine.parseColumn(m.group(1))};
        }

        private Object readCell(int r, int c, String prop) throws MacroException {
            switch (prop) {
                case "value":
                case "text": {
                    Object v = model.getComputedValue(r, c);
                    return v == null ? "" : v;
                }
                case "formula":
                    return model.getInput(r, c);
                default:
                    throw err("Unsupported property: " + prop);
            }
        }

        private Object callFunction(String name, List<Object> a) throws MacroException {
            switch (name) {
                case "int":
                    need(a, 1);
                    return Math.floor(toNum(a.get(0)));
                case "abs":
                    need(a, 1);
                    return Math.abs(toNum(a.get(0)));
                case "sqr": {
                    need(a, 1);
                    double x = toNum(a.get(0));
                    if (x < 0) throw err("Invalid procedure call: Sqr of a negative number");
                    return Math.sqrt(x);
                }
                case "rnd":
                    return random.nextDouble();
                case "round": {
                    if (a.isEmpty() || a.size() > 2) throw err("Wrong number of arguments for Round");
                    int digits = a.size() == 2 ? (int) toNum(a.get(1)) : 0;
                    return new BigDecimal(toNum(a.get(0))).setScale(digits, RoundingMode.HALF_UP).doubleValue();
                }
                case "len":
                    need(a, 1);
                    return (double) str(a.get(0)).length();
                case "ucase":
                    need(a, 1);
                    return str(a.get(0)).toUpperCase(Locale.ROOT);
                case "lcase":
                    need(a, 1);
                    return str(a.get(0)).toLowerCase(Locale.ROOT);
                case "str":
                    need(a, 1);
                    return str(a.get(0));
                case "val": {
                    need(a, 1);
                    Matcher m = LEADING_NUMBER.matcher(str(a.get(0)));
                    return m.find() ? Double.parseDouble(m.group().trim()) : 0.0;
                }
                case "mid": {
                    if (a.size() < 2 || a.size() > 3) throw err("Wrong number of arguments for Mid");
                    String s = str(a.get(0));
                    int st = Math.max(1, (int) toNum(a.get(1))) - 1;
                    if (st >= s.length()) return "";
                    int end = a.size() == 3 ? st + Math.max(0, (int) toNum(a.get(2))) : s.length();
                    return s.substring(st, Math.min(s.length(), end));
                }
                case "rgb": {
                    need(a, 3);
                    int r = (int) toNum(a.get(0)) & 255;
                    int g = (int) toNum(a.get(1)) & 255;
                    int b = (int) toNum(a.get(2)) & 255;
                    return (double) (r | (g << 8) | (b << 16));
                }
                default:
                    throw err("Unknown function: " + name);
            }
        }

        private void need(List<Object> a, int n) throws MacroException {
            if (a.size() != n) throw err("Wrong number of arguments (expected " + n + ")");
        }
    }
}
