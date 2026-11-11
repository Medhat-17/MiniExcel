import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.security.GeneralSecurityException;
import java.security.SecureRandom;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.zip.GZIPInputStream;
import java.util.zip.GZIPOutputStream;
import javax.crypto.AEADBadTagException;
import javax.crypto.Cipher;
import javax.crypto.SecretKey;
import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.GCMParameterSpec;
import javax.crypto.spec.PBEKeySpec;
import javax.crypto.spec.SecretKeySpec;

/**
 * File layer: saves and loads whole workbooks in the native ".spx" format.
 *
 * Layout: a 4-byte magic number followed by a GZIP-compressed binary payload
 * (sheets, cell inputs, styles, protection hashes, pictures). When an open
 * password is set the payload is instead encrypted with AES-256-GCM using a key
 * derived from the password with PBKDF2-HMAC-SHA256; a wrong password makes the
 * authentication tag fail, so it can be told apart from a good file.
 */
public final class WorkbookIO {

    public static final String EXTENSION = "spx";

    private static final int MAGIC_PLAIN = 0x53505831;     // "SPX1"
    private static final int MAGIC_ENCRYPTED = 0x53505832; // "SPX2"
    private static final int FORMAT_VERSION = 1;
    private static final int KDF_ITERATIONS = 120_000;
    private static final int SALT_BYTES = 16;
    private static final int IV_BYTES = 12;
    private static final int MAX_FILE_BYTES = 512 * 1024 * 1024;
    private static final int MAX_STRING_BYTES = 8 * 1024 * 1024;
    private static final int MAX_IMAGE_BYTES = 64 * 1024 * 1024;

    private WorkbookIO() {
    }

    /** One sheet to save / one sheet that was loaded. */
    public static final class SheetData {
        public final String name;
        public final SpreadsheetModel model;

        public SheetData(String name, SpreadsheetModel model) {
            this.name = name;
            this.model = model;
        }
    }

    /** Result of {@link #load}. */
    public static final class WorkbookData {
        public final List<SheetData> sheets;
        public final int activeIndex;

        WorkbookData(List<SheetData> sheets, int activeIndex) {
            this.sheets = sheets;
            this.activeIndex = activeIndex;
        }
    }

    /** The file is encrypted and no password was supplied. */
    public static class PasswordRequiredException extends IOException {
        private static final long serialVersionUID = 1L;

        public PasswordRequiredException() {
            super("This workbook is protected with an open password.");
        }
    }

    /** The supplied open password is wrong (or the encrypted file is damaged). */
    public static class BadPasswordException extends IOException {
        private static final long serialVersionUID = 1L;

        public BadPasswordException() {
            super("Incorrect password (or the file is damaged).");
        }
    }

    // ==================================================================
    // Save
    // ==================================================================

    /**
     * Writes the workbook. A non-empty {@code password} encrypts the file.
     * The data goes to a temporary file first so an existing workbook is never left half-written.
     */
    public static void save(File file, List<SheetData> sheets, int activeIndex, char[] password)
            throws IOException {
        byte[] payload = serialize(sheets, activeIndex);
        byte[] out;
        if (password == null || password.length == 0) {
            out = new byte[4 + payload.length];
            putInt(out, 0, MAGIC_PLAIN);
            System.arraycopy(payload, 0, out, 4, payload.length);
        } else {
            out = encrypt(payload, password);
        }
        Path target = file.toPath().toAbsolutePath();
        Path tmp = target.resolveSibling(target.getFileName() + ".tmp");
        Files.write(tmp, out);
        Files.move(tmp, target, StandardCopyOption.REPLACE_EXISTING);
    }

    private static byte[] serialize(List<SheetData> sheets, int activeIndex) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try (DataOutputStream out = new DataOutputStream(new GZIPOutputStream(bos))) {
            out.writeInt(FORMAT_VERSION);
            out.writeInt(activeIndex);
            out.writeInt(sheets.size());
            for (SheetData s : sheets) writeSheet(out, s);
        }
        return bos.toByteArray();
    }

    private static void writeSheet(DataOutputStream out, SheetData s) throws IOException {
        SpreadsheetModel m = s.model;
        writeString(out, s.name);

        SheetProtection p = m.getProtection();
        byte[] salt = p.getSalt();
        byte[] hash = p.getHash();
        out.writeBoolean(p.isProtected());
        out.writeBoolean(salt != null && hash != null);
        if (salt != null && hash != null) {
            writeBytes(out, salt);
            writeBytes(out, hash);
        }

        int rows = m.getRowCount();
        int cols = m.getColumnCount();
        out.writeInt(rows);
        out.writeInt(cols);

        int cells = 0;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                if (!m.getInput(r, c).isEmpty()) cells++;
        out.writeInt(cells);
        for (int r = 0; r < rows; r++) {
            for (int c = 0; c < cols; c++) {
                String text = m.getInput(r, c);
                if (text.isEmpty()) continue;
                out.writeInt(r);
                out.writeInt(c);
                writeString(out, text);
            }
        }

        int styled = 0;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                if (m.getStyle(r, c) != null) styled++;
        out.writeInt(styled);
        for (int r = 0; r < rows; r++) {
            for (int c = 0; c < cols; c++) {
                CellStyle st = m.getStyle(r, c);
                if (st == null) continue;
                out.writeInt(r);
                out.writeInt(c);
                int flags = (st.bold ? 1 : 0) | (st.italic ? 2 : 0)
                        | (st.background != null ? 4 : 0) | (st.foreground != null ? 8 : 0);
                out.writeByte(flags);
                out.writeInt(st.alignment);
                out.writeInt(st.fontSize);
                if (st.background != null) out.writeInt(st.background.getRGB());
                if (st.foreground != null) out.writeInt(st.foreground.getRGB());
            }
        }

        List<SheetImage> images = m.getImages();
        out.writeInt(images.size());
        for (SheetImage im : images) {
            out.writeInt(im.x);
            out.writeInt(im.y);
            out.writeInt(im.width);
            out.writeInt(im.height);
            writeBytes(out, im.getPngBytes());
        }
    }

    // ==================================================================
    // Load
    // ==================================================================

    /**
     * Reads a workbook.
     * @param password the open password, or null if none is known yet
     * @throws PasswordRequiredException if the file is encrypted and {@code password} is null/empty
     * @throws BadPasswordException      if the password does not decrypt the file
     */
    public static WorkbookData load(File file, char[] password) throws IOException {
        if (file.length() > MAX_FILE_BYTES) throw new IOException("File is too large");
        byte[] bytes = Files.readAllBytes(file.toPath());
        if (bytes.length < 8) throw new IOException("Not a Spreadsheet Pro workbook");
        int magic = getInt(bytes, 0);
        byte[] payload;
        if (magic == MAGIC_PLAIN) {
            payload = Arrays.copyOfRange(bytes, 4, bytes.length);
        } else if (magic == MAGIC_ENCRYPTED) {
            payload = decrypt(bytes, password);
        } else {
            throw new IOException("Not a Spreadsheet Pro workbook");
        }
        try {
            return parse(payload);
        } catch (RuntimeException e) {
            throw new IOException("The workbook file is damaged: " + e.getMessage(), e);
        }
    }

    private static WorkbookData parse(byte[] payload) throws IOException {
        try (DataInputStream in = new DataInputStream(new GZIPInputStream(new ByteArrayInputStream(payload)))) {
            int version = in.readInt();
            if (version < 1 || version > FORMAT_VERSION) {
                throw new IOException("This workbook was made by a newer version (format " + version + ")");
            }
            int active = in.readInt();
            int count = in.readInt();
            if (count < 1 || count > 255) throw new IOException("Invalid sheet count");
            List<SheetData> sheets = new ArrayList<>();
            for (int i = 0; i < count; i++) sheets.add(readSheet(in));
            // Read to the end of the stream: GZIP only verifies its CRC at the trailer,
            // so this is what detects a damaged file.
            byte[] sink = new byte[4096];
            while (in.read(sink) != -1) {
                // discard
            }
            return new WorkbookData(sheets, Math.max(0, Math.min(active, count - 1)));
        }
    }

    private static SheetData readSheet(DataInputStream in) throws IOException {
        String name = readString(in);
        boolean prot = in.readBoolean();
        boolean hasPassword = in.readBoolean();
        byte[] salt = null;
        byte[] hash = null;
        if (hasPassword) {
            salt = readBytes(in, 1024);
            hash = readBytes(in, 1024);
        }
        int rows = in.readInt();
        int cols = in.readInt();
        if (rows < 1 || rows > 10_000 || cols < 1 || cols > 702) throw new IOException("Invalid sheet size");

        SpreadsheetModel m = new SpreadsheetModel(rows, cols);
        m.beginBatch();
        try {
            int cells = in.readInt();
            if (cells < 0 || cells > (long) rows * cols) throw new IOException("Invalid cell count");
            for (int i = 0; i < cells; i++) {
                int r = in.readInt();
                int c = in.readInt();
                String text = readString(in);
                if (m.inBounds(r, c)) m.setInput(r, c, text);
            }

            int styled = in.readInt();
            if (styled < 0 || styled > (long) rows * cols) throw new IOException("Invalid style count");
            for (int i = 0; i < styled; i++) {
                final int r = in.readInt();
                final int c = in.readInt();
                final int flags = in.readUnsignedByte();
                final int align = in.readInt();
                final int size = in.readInt();
                final Color bg = (flags & 4) != 0 ? new Color(in.readInt(), true) : null;
                final Color fg = (flags & 8) != 0 ? new Color(in.readInt(), true) : null;
                m.updateStyle(r, c, st -> {
                    st.bold = (flags & 1) != 0;
                    st.italic = (flags & 2) != 0;
                    st.alignment = align;
                    st.fontSize = size;
                    st.background = bg;
                    st.foreground = fg;
                });
            }

            int pictures = in.readInt();
            if (pictures < 0 || pictures > 1000) throw new IOException("Invalid picture count");
            for (int i = 0; i < pictures; i++) {
                int x = in.readInt();
                int y = in.readInt();
                int w = in.readInt();
                int h = in.readInt();
                byte[] data = readBytes(in, MAX_IMAGE_BYTES);
                try {
                    m.addImage(SheetImage.fromPng(data, x, y, w, h));
                } catch (IOException skip) {
                    // an unreadable picture is dropped; the rest of the sheet still loads
                }
            }
        } finally {
            m.endBatch();
        }
        m.getProtection().restore(prot, salt, hash); // last, so loading is not blocked by it
        return new SheetData(name, m);
    }

    // ==================================================================
    // Encryption
    // ==================================================================

    private static byte[] encrypt(byte[] payload, char[] password) throws IOException {
        try {
            SecureRandom rnd = new SecureRandom();
            byte[] salt = new byte[SALT_BYTES];
            byte[] iv = new byte[IV_BYTES];
            rnd.nextBytes(salt);
            rnd.nextBytes(iv);
            Cipher cipher = Cipher.getInstance("AES/GCM/NoPadding");
            cipher.init(Cipher.ENCRYPT_MODE, deriveKey(password, salt), new GCMParameterSpec(128, iv));
            byte[] encrypted = cipher.doFinal(payload);
            byte[] out = new byte[4 + SALT_BYTES + IV_BYTES + encrypted.length];
            putInt(out, 0, MAGIC_ENCRYPTED);
            System.arraycopy(salt, 0, out, 4, SALT_BYTES);
            System.arraycopy(iv, 0, out, 4 + SALT_BYTES, IV_BYTES);
            System.arraycopy(encrypted, 0, out, 4 + SALT_BYTES + IV_BYTES, encrypted.length);
            return out;
        } catch (GeneralSecurityException e) {
            throw new IOException("Encryption failed: " + e.getMessage(), e);
        }
    }

    private static byte[] decrypt(byte[] file, char[] password) throws IOException {
        if (password == null || password.length == 0) throw new PasswordRequiredException();
        int header = 4 + SALT_BYTES + IV_BYTES;
        if (file.length < header + 16) throw new IOException("The workbook file is truncated");
        byte[] salt = Arrays.copyOfRange(file, 4, 4 + SALT_BYTES);
        byte[] iv = Arrays.copyOfRange(file, 4 + SALT_BYTES, header);
        try {
            Cipher cipher = Cipher.getInstance("AES/GCM/NoPadding");
            cipher.init(Cipher.DECRYPT_MODE, deriveKey(password, salt), new GCMParameterSpec(128, iv));
            return cipher.doFinal(file, header, file.length - header);
        } catch (AEADBadTagException e) {
            throw new BadPasswordException();
        } catch (GeneralSecurityException e) {
            throw new IOException("Decryption failed: " + e.getMessage(), e);
        }
    }

    private static SecretKey deriveKey(char[] password, byte[] salt) throws GeneralSecurityException {
        PBEKeySpec spec = new PBEKeySpec(password, salt, KDF_ITERATIONS, 256);
        try {
            byte[] key = SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256").generateSecret(spec).getEncoded();
            return new SecretKeySpec(key, "AES");
        } finally {
            spec.clearPassword();
        }
    }

    // ==================================================================
    // Small helpers
    // ==================================================================

    private static void putInt(byte[] b, int off, int v) {
        b[off] = (byte) (v >>> 24);
        b[off + 1] = (byte) (v >>> 16);
        b[off + 2] = (byte) (v >>> 8);
        b[off + 3] = (byte) v;
    }

    private static int getInt(byte[] b, int off) {
        return ((b[off] & 0xFF) << 24) | ((b[off + 1] & 0xFF) << 16) | ((b[off + 2] & 0xFF) << 8) | (b[off + 3] & 0xFF);
    }

    private static void writeString(DataOutputStream out, String s) throws IOException {
        writeBytes(out, s.getBytes(StandardCharsets.UTF_8));
    }

    private static String readString(DataInputStream in) throws IOException {
        return new String(readBytes(in, MAX_STRING_BYTES), StandardCharsets.UTF_8);
    }

    private static void writeBytes(DataOutputStream out, byte[] b) throws IOException {
        out.writeInt(b.length);
        out.write(b);
    }

    private static byte[] readBytes(DataInputStream in, int max) throws IOException {
        int len = in.readInt();
        if (len < 0 || len > max) throw new IOException("Invalid data length");
        byte[] b = new byte[len];
        in.readFully(b);
        return b;
    }
}
