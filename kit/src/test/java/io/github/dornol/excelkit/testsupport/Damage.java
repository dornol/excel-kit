package io.github.dornol.excelkit.testsupport;

import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * Ways of damaging a valid file for the tests that hand a reader or a template writer something a client could have
 * sent: bytes cut off, flipped, overwritten, zeroed, duplicated or removed, and, for a workbook, the XML inside the
 * archive damaged and zipped again so that it reaches the parsers. Each method returns what it did, to put in a
 * failure message so that the case can be repeated from its seed.
 */
public final class Damage {
    private Damage() {
    }

    public static String raw(Random r, byte[] data, byte[][] result) {
        byte[] b = data.clone();
        int n = b.length;
        switch (r.nextInt(8)) {
            case 0 -> {
                int len = r.nextInt(4) == 0 ? r.nextInt(Math.min(n, 64)) : r.nextInt(n);
                result[0] = java.util.Arrays.copyOf(b, len);
                return "truncated to " + len + " of " + n;
            }
            case 1 -> {
                int flips = 1 + r.nextInt(16);
                for (int i = 0; i < flips; i++) b[r.nextInt(n)] ^= (byte) (1 << r.nextInt(8));
                result[0] = b;
                return flips + " bit flips";
            }
            case 2 -> {
                int at = r.nextInt(n), len = Math.min(1 + r.nextInt(64), n - at);
                byte[] garbage = new byte[len];
                r.nextBytes(garbage);
                System.arraycopy(garbage, 0, b, at, len);
                result[0] = b;
                return "overwrote " + len + " bytes at " + at;
            }
            case 3 -> {
                int at = r.nextInt(n), len = Math.min(1 + r.nextInt(256), n - at);
                java.util.Arrays.fill(b, at, at + len, (byte) 0);
                result[0] = b;
                return "zeroed " + len + " bytes at " + at;
            }
            case 4 -> {
                int at = r.nextInt(n), len = Math.min(1 + r.nextInt(512), n - at), to = r.nextInt(n);
                byte[] out = new byte[n + len];
                System.arraycopy(b, 0, out, 0, to);
                System.arraycopy(b, at, out, to, len);
                System.arraycopy(b, to, out, to + len, n - to);
                result[0] = out;
                return "duplicated " + len + " bytes from " + at + " at " + to;
            }
            case 5 -> {
                int at = r.nextInt(n), len = Math.min(1 + r.nextInt(512), n - at);
                byte[] out = new byte[n - len];
                System.arraycopy(b, 0, out, 0, at);
                System.arraycopy(b, at + len, out, at, n - at - len);
                result[0] = out;
                return "removed " + len + " bytes at " + at;
            }
            case 6 -> {
                // the end of a ZIP holds its central directory: damage there
                int span = Math.min(n, 120), at = n - 1 - r.nextInt(span);
                b[at] = (byte) r.nextInt(256);
                if (at > 0) b[at - 1] = (byte) r.nextInt(256);
                result[0] = b;
                return "damaged the tail at " + at;
            }
            default -> {
                int at = r.nextInt(n), len = Math.min(1 + r.nextInt(8), n - at);
                for (int i = 0; i < len; i++) b[at + i] = (byte) (r.nextBoolean() ? 0xFF : 0x00);
                result[0] = b;
                return "set " + len + " bytes to 00/FF at " + at;
            }
        }
    }

    /** Damages the XML of one part of the archive and zips it again, so the archive itself stays valid. */
    public static String insideArchive(Random r, byte[] xlsx, byte[][] result) throws Exception {
        Path tmp = Files.createTempFile("fuzz", ".xlsx");
        try {
            Files.write(tmp, xlsx);
            Map<String, byte[]> parts = new LinkedHashMap<>();
            try (ZipFile zip = new ZipFile(tmp.toFile())) {
                for (Enumeration<? extends ZipEntry> e = zip.entries(); e.hasMoreElements(); ) {
                    ZipEntry entry = e.nextElement();
                    parts.put(entry.getName(), zip.getInputStream(entry).readAllBytes());
                }
            }
            List<String> names = new ArrayList<>(parts.keySet());
            String target = names.get(r.nextInt(names.size()));
            byte[] content = parts.get(target);
            if (content.length == 0) {
                result[0] = xlsx;
                return "left " + target + " (empty)";
            }
            byte[] changed = content.clone();
            String how;
            switch (r.nextInt(5)) {
                case 0 -> {
                    changed = java.util.Arrays.copyOf(content, r.nextInt(content.length));
                    how = "cut to " + changed.length;
                }
                case 1 -> {
                    for (int i = 0, n = 1 + r.nextInt(8); i < n; i++) changed[r.nextInt(changed.length)] = (byte) r.nextInt(256);
                    how = "random bytes";
                }
                case 2 -> {
                    // break the markup: turn a '<' or '>' into something else
                    for (int i = 0, n = 1 + r.nextInt(4); i < n; i++) {
                        int at = r.nextInt(changed.length);
                        for (int k = 0; k < changed.length; k++) {
                            int p = (at + k) % changed.length;
                            if (changed[p] == '<' || changed[p] == '>' || changed[p] == '"') {
                                changed[p] = (byte) "&;/ ".charAt(r.nextInt(4));
                                break;
                            }
                        }
                    }
                    how = "broken markup";
                }
                case 3 -> {
                    // swap a number or an attribute value for something hostile
                    String text = new String(content, java.nio.charset.StandardCharsets.UTF_8);
                    String[] hostile = {"-1", "9999999999", "NaN", "1e999", "", "0", "2147483648", "A0", "XFE1", "&#0;", "&unknown;"};
                    java.util.regex.Matcher m = java.util.regex.Pattern.compile("\"(\\d+)\"|>(\\d+)<").matcher(text);
                    List<int[]> spots = new ArrayList<>();
                    while (m.find()) spots.add(new int[]{m.start(), m.end()});
                    if (spots.isEmpty()) {
                        how = "no numbers to replace";
                    } else {
                        int[] spot = spots.get(r.nextInt(spots.size()));
                        String value = hostile[r.nextInt(hostile.length)];
                        String old = text.substring(spot[0], spot[1]);
                        String replacement = old.startsWith("\"") ? "\"" + value + "\"" : ">" + value + "<";
                        changed = (text.substring(0, spot[0]) + replacement + text.substring(spot[1])).getBytes(java.nio.charset.StandardCharsets.UTF_8);
                        how = "number -> " + value;
                    }
                }
                default -> {
                    int len = Math.min(1 + r.nextInt(200), content.length);
                    int from = r.nextInt(content.length - len + 1);
                    byte[] out = new byte[content.length + len];
                    System.arraycopy(content, 0, out, 0, from + len);
                    System.arraycopy(content, from, out, from + len, content.length - from);
                    changed = out;
                    how = "repeated " + len + " bytes";
                }
            }
            parts.put(target, changed);
            // sometimes drop the part altogether
            if (r.nextInt(12) == 0) {
                parts.remove(target);
                how = "removed the part";
            }
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ZipOutputStream zip = new ZipOutputStream(out)) {
                for (var e : parts.entrySet()) {
                    zip.putNextEntry(new ZipEntry(e.getKey()));
                    zip.write(e.getValue());
                    zip.closeEntry();
                }
            }
            result[0] = out.toByteArray();
            return target + ": " + how;
        } finally {
            Files.deleteIfExists(tmp);
        }
    }
}
