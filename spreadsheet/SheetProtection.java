import java.security.GeneralSecurityException;
import java.security.MessageDigest;
import java.security.SecureRandom;
import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.PBEKeySpec;

/**
 * Protection state of one worksheet.
 *
 * A protected sheet cannot be edited, formatted, cleared, changed by a macro or
 * have its pictures moved. The optional password is never stored: only a random
 * salt and a PBKDF2-HMAC-SHA256 hash of it.
 *
 * Note: this stops changes inside the application (like Excel's sheet protection).
 * To keep the file contents private use File &gt; Set Open Password, which encrypts
 * the saved workbook.
 */
public final class SheetProtection {

    private static final int ITERATIONS = 40_000;
    private static final int SALT_BYTES = 16;
    private static final int KEY_BITS = 256;

    private boolean on;
    private byte[] salt;
    private byte[] hash;

    public boolean isProtected() {
        return on;
    }

    /** True when unprotecting requires a password. */
    public boolean hasPassword() {
        return on && hash != null;
    }

    /** Turns protection on. A null or empty password protects without a password. */
    public void protect(char[] password) {
        on = true;
        if (password == null || password.length == 0) {
            salt = null;
            hash = null;
            return;
        }
        salt = new byte[SALT_BYTES];
        new SecureRandom().nextBytes(salt);
        hash = derive(password, salt);
    }

    /** @return true if the password was right (or none was needed) and protection is now off. */
    public boolean unprotect(char[] password) {
        if (!on) return true;
        if (!verify(password)) return false;
        on = false;
        salt = null;
        hash = null;
        return true;
    }

    public boolean verify(char[] password) {
        if (hash == null) return true;
        if (password == null) return false;
        return MessageDigest.isEqual(hash, derive(password, salt));
    }

    // ---- persistence helpers (used by WorkbookIO) ----

    public byte[] getSalt() {
        return salt == null ? null : salt.clone();
    }

    public byte[] getHash() {
        return hash == null ? null : hash.clone();
    }

    public void restore(boolean protectedFlag, byte[] savedSalt, byte[] savedHash) {
        on = protectedFlag;
        if (protectedFlag && savedSalt != null && savedHash != null) {
            salt = savedSalt.clone();
            hash = savedHash.clone();
        } else {
            salt = null;
            hash = null;
        }
    }

    private static byte[] derive(char[] password, byte[] salt) {
        PBEKeySpec spec = new PBEKeySpec(password, salt, ITERATIONS, KEY_BITS);
        try {
            return SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256").generateSecret(spec).getEncoded();
        } catch (GeneralSecurityException e) {
            throw new IllegalStateException("PBKDF2 not available", e);
        } finally {
            spec.clearPassword();
        }
    }
}
