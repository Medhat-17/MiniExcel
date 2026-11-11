import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Calculation layer: a recursive-descent evaluator for spreadsheet formulas.
 *
 * Grammar (lowest to highest precedence):
 * <pre>
 *   expr    := term   (('+' | '-') term)*
 *   term    := power  (('*' | '/') power)*
 *   power   := unary  ('^' power)?
 *   unary   := ('-' | '+') unary | postfix
 *   postfix := primary '%'*
 *   primary := NUMBER | CELLREF | FUNC '(' args ')' | '(' expr ')'
 * </pre>
 * Functions: SUM, AVERAGE, MIN, MAX, COUNT (accept ranges and lists),
 * plus ABS, SQRT, ROUND.
 *
 * The engine never throws: {@link #evaluate(String)} returns either a
 * {@link Double} or one of the error strings #DIV/0!, #SYNTAX!, #NAME?, #ERROR!.
 */
public class FormulaEngine {

    public static final String ERR_DIV0 = "#DIV/0!";
    public static final String ERR_SYNTAX = "#SYNTAX!";
    public static final String ERR_NAME = "#NAME?";
    public static final String ERR_GENERIC = "#ERROR!";

    /** Supplies cell values to the engine (implemented by the model). */
    public interface CellResolver {
        /** @return Double, String (text or error code) or null when the cell is empty. */
        Object resolve(int row, int col);
    }

    /** Internal control-flow exception carrying a spreadsheet error code. */
    public static class FormulaException extends Exception {
        private final String code;

        public FormulaException(String code) {
            super(code, null, false, false); // no stack trace: cheap to throw
            this.code = code;
        }

        public String getCode() {
            return code;
        }
    }

    private static final Pattern RANGE_ARG =
            Pattern.compile("([A-Za-z]+)(\\d+):([A-Za-z]+)(\\d+)\\s*(?=[,)])");
    private static final Pattern CELL_ARG =
            Pattern.compile("([A-Za-z]+)(\\d+)\\s*(?=[,)])");
    private static final Set<String> FUNCTIONS = new HashSet<>(Arrays.asList(
            "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "ABS", "SQRT", "ROUND"));

    private final CellResolver resolver;
    private final int maxRows;
    private final int maxCols;

    public FormulaEngine(CellResolver resolver, int maxRows, int maxCols) {
        this.resolver = resolver;
        this.maxRows = maxRows;
        this.maxCols = maxCols;
    }

    // ------------------------------------------------------------------
    // Public helpers shared with the other layers
    // ------------------------------------------------------------------

    /** 0 -> "A", 25 -> "Z", 26 -> "AA". */
    public static String columnLabel(int index) {
        StringBuilder sb = new StringBuilder();
        int n = index + 1;
        while (n > 0) {
            n--;
            sb.insert(0, (char) ('A' + n % 26));
            n /= 26;
        }
        return sb.toString();
    }

    /** "A" -> 0, "Z" -> 25, "AA" -> 26 (case-insensitive). */
    public static int parseColumn(String letters) {
        int value = 0;
        for (char ch : letters.toUpperCase(Locale.ROOT).toCharArray()) {
            value = value * 26 + (ch - 'A' + 1);
        }
        return value - 1;
    }

    /** (0,0) -> "A1". */
    public static String cellRef(int row, int col) {
        return columnLabel(col) + (row + 1);
    }

    public static boolean isErrorCode(String s) {
        return ERR_DIV0.equals(s) || ERR_SYNTAX.equals(s) || ERR_NAME.equals(s) || ERR_GENERIC.equals(s);
    }

    /** Formats a number the way a spreadsheet displays it (no trailing zeros, no ".0"). */
    public static String formatNumber(double d) {
        if (Double.isNaN(d) || Double.isInfinite(d)) return ERR_GENERIC;
        if (d == Math.rint(d) && Math.abs(d) < 1e15) return Long.toString((long) d);
        BigDecimal bd = new BigDecimal(d).round(new MathContext(12)).stripTrailingZeros();
        return bd.toPlainString();
    }

    // ------------------------------------------------------------------
    // Evaluation entry point
    // ------------------------------------------------------------------

    /**
     * Evaluates a formula such as "=A1+B2*5" or "=SUM(A1:A10)/COUNT(A1:A10)".
     * @return a Double, or an error String
     */
    public Object evaluate(String formula) {
        try {
            String body = formula.trim();
            if (body.startsWith("=")) body = body.substring(1);
            double v = new Parser(body).parseAll();
            if (Double.isNaN(v) || Double.isInfinite(v)) return ERR_GENERIC;
            return v;
        } catch (FormulaException e) {
            return e.getCode();
        } catch (NumberFormatException e) {
            return ERR_SYNTAX;
        } catch (StackOverflowError e) {
            return ERR_GENERIC;
        } catch (RuntimeException e) {
            return ERR_GENERIC;
        }
    }

    private static boolean isAsciiLetter(char c) {
        return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
    }

    // ------------------------------------------------------------------
    // Parser
    // ------------------------------------------------------------------

    private class Parser {
        private final String s;
        private int pos = 0;

        Parser(String s) {
            this.s = s;
        }

        double parseAll() throws FormulaException {
            skipWs();
            if (pos >= s.length()) throw syntax();
            double v = expr();
            skipWs();
            if (pos < s.length()) throw syntax();
            return v;
        }

        private FormulaException syntax() {
            return new FormulaException(ERR_SYNTAX);
        }

        private void skipWs() {
            while (pos < s.length() && Character.isWhitespace(s.charAt(pos))) pos++;
        }

        private boolean match(char c) {
            skipWs();
            if (pos < s.length() && s.charAt(pos) == c) {
                pos++;
                return true;
            }
            return false;
        }

        private double expr() throws FormulaException {
            double v = term();
            while (true) {
                if (match('+')) v += term();
                else if (match('-')) v -= term();
                else return v;
            }
        }

        private double term() throws FormulaException {
            double v = power();
            while (true) {
                if (match('*')) {
                    v *= power();
                } else if (match('/')) {
                    double d = power();
                    if (d == 0) throw new FormulaException(ERR_DIV0);
                    v /= d;
                } else {
                    return v;
                }
            }
        }

        private double power() throws FormulaException {
            double base = unary();
            if (match('^')) return Math.pow(base, power());
            return base;
        }

        private double unary() throws FormulaException {
            if (match('-')) return -unary();
            if (match('+')) return unary();
            return postfix();
        }

        private double postfix() throws FormulaException {
            double v = primary();
            while (match('%')) v /= 100.0;
            return v;
        }

        private double primary() throws FormulaException {
            skipWs();
            if (pos >= s.length()) throw syntax();
            char c = s.charAt(pos);
            if (c == '(') {
                pos++;
                double v = expr();
                if (!match(')')) throw syntax();
                return v;
            }
            if (Character.isDigit(c) || c == '.') return number();
            if (isAsciiLetter(c)) return identifier();
            throw syntax();
        }

        private double number() {
            int start = pos;
            while (pos < s.length() && (Character.isDigit(s.charAt(pos)) || s.charAt(pos) == '.')) pos++;
            if (pos < s.length() && (s.charAt(pos) == 'e' || s.charAt(pos) == 'E')) {
                int save = pos;
                pos++;
                if (pos < s.length() && (s.charAt(pos) == '+' || s.charAt(pos) == '-')) pos++;
                if (pos < s.length() && Character.isDigit(s.charAt(pos))) {
                    while (pos < s.length() && Character.isDigit(s.charAt(pos))) pos++;
                } else {
                    pos = save; // not an exponent after all
                }
            }
            return Double.parseDouble(s.substring(start, pos)); // NumberFormatException -> #SYNTAX!
        }

        /** Function call or single cell reference. */
        private double identifier() throws FormulaException {
            int start = pos;
            while (pos < s.length() && isAsciiLetter(s.charAt(pos))) pos++;
            String letters = s.substring(start, pos);

            // Function call?  NAME '('
            int save = pos;
            skipWs();
            if (pos < s.length() && s.charAt(pos) == '(') {
                String name = letters.toUpperCase(Locale.ROOT);
                if (!FUNCTIONS.contains(name)) throw new FormulaException(ERR_NAME);
                pos++; // consume '('
                return callFunction(name, parseArgs());
            }
            pos = save;

            // Cell reference?  LETTERS DIGITS
            int digitStart = pos;
            while (pos < s.length() && Character.isDigit(s.charAt(pos))) pos++;
            if (pos == digitStart || letters.length() > 3) throw new FormulaException(ERR_NAME);
            int row = Integer.parseInt(s.substring(digitStart, pos)) - 1;
            int col = parseColumn(letters);
            skipWs();
            if (pos < s.length() && s.charAt(pos) == ':') throw syntax(); // range outside a function
            return cellNumber(row, col);
        }

        private double cellNumber(int row, int col) throws FormulaException {
            if (row < 0 || col < 0 || row >= maxRows || col >= maxCols) throw new FormulaException(ERR_GENERIC);
            Object v = resolver.resolve(row, col);
            if (v == null) return 0;
            if (v instanceof Double) return (Double) v;
            String str = v.toString();
            if (isErrorCode(str)) throw new FormulaException(str); // propagate upstream errors
            throw new FormulaException(ERR_GENERIC);               // text used as a number
        }

        /** Parses "arg, arg, ...)" after the opening parenthesis. Each arg becomes a double[]. */
        private List<double[]> parseArgs() throws FormulaException {
            List<double[]> args = new ArrayList<>();
            skipWs();
            if (pos < s.length() && s.charAt(pos) == ')') {
                pos++;
                return args;
            }
            while (true) {
                skipWs();
                Matcher range = RANGE_ARG.matcher(s);
                range.region(pos, s.length());
                if (range.lookingAt()) {
                    args.add(collectRange(range.group(1), range.group(2), range.group(3), range.group(4)));
                    pos = range.end();
                } else {
                    Matcher cell = CELL_ARG.matcher(s);
                    cell.region(pos, s.length());
                    if (cell.lookingAt()) {
                        args.add(collectRange(cell.group(1), cell.group(2), cell.group(1), cell.group(2)));
                        pos = cell.end();
                    } else {
                        args.add(new double[]{expr()});
                    }
                }
                skipWs();
                if (pos < s.length() && s.charAt(pos) == ',') {
                    pos++;
                } else if (pos < s.length() && s.charAt(pos) == ')') {
                    pos++;
                    return args;
                } else {
                    throw syntax();
                }
            }
        }

        /** Collects the numeric cells of a rectangular range (empty and text cells are skipped). */
        private double[] collectRange(String l1, String d1, String l2, String d2) throws FormulaException {
            int c1 = parseColumn(l1), c2 = parseColumn(l2);
            int r1 = Integer.parseInt(d1) - 1, r2 = Integer.parseInt(d2) - 1;
            int rowLo = Math.min(r1, r2), rowHi = Math.max(r1, r2);
            int colLo = Math.min(c1, c2), colHi = Math.max(c1, c2);
            if (rowLo < 0 || colLo < 0 || rowHi >= maxRows || colHi >= maxCols) {
                throw new FormulaException(ERR_GENERIC);
            }
            List<Double> values = new ArrayList<>();
            for (int r = rowLo; r <= rowHi; r++) {
                for (int c = colLo; c <= colHi; c++) {
                    Object v = resolver.resolve(r, c);
                    if (v instanceof Double) {
                        values.add((Double) v);
                    } else if (v instanceof String && isErrorCode((String) v)) {
                        throw new FormulaException((String) v);
                    }
                }
            }
            double[] out = new double[values.size()];
            for (int i = 0; i < out.length; i++) out[i] = values.get(i);
            return out;
        }

        private double single(List<double[]> args) throws FormulaException {
            if (args.size() != 1 || args.get(0).length != 1) throw new FormulaException(ERR_GENERIC);
            return args.get(0)[0];
        }

        private double callFunction(String name, List<double[]> args) throws FormulaException {
            double sum = 0;
            int n = 0;
            double min = Double.POSITIVE_INFINITY, max = Double.NEGATIVE_INFINITY;
            for (double[] a : args) {
                for (double v : a) {
                    sum += v;
                    n++;
                    if (v < min) min = v;
                    if (v > max) max = v;
                }
            }
            switch (name) {
                case "SUM":
                    return sum;
                case "AVERAGE":
                    if (n == 0) throw new FormulaException(ERR_DIV0);
                    return sum / n;
                case "MIN":
                    return n == 0 ? 0 : min;
                case "MAX":
                    return n == 0 ? 0 : max;
                case "COUNT":
                    return n;
                case "ABS":
                    return Math.abs(single(args));
                case "SQRT": {
                    double x = single(args);
                    if (x < 0) throw new FormulaException(ERR_GENERIC);
                    return Math.sqrt(x);
                }
                case "ROUND": {
                    if (args.isEmpty() || args.size() > 2) throw new FormulaException(ERR_GENERIC);
                    for (double[] a : args) if (a.length != 1) throw new FormulaException(ERR_GENERIC);
                    double x = args.get(0)[0];
                    int digits = args.size() == 2 ? (int) args.get(1)[0] : 0;
                    return new BigDecimal(x).setScale(digits, RoundingMode.HALF_UP).doubleValue();
                }
                default:
                    throw new FormulaException(ERR_NAME);
            }
        }
    }
}
