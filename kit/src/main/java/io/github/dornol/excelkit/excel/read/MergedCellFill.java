package io.github.dornol.excelkit.excel.read;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.XMLHelper;
import org.jspecify.annotations.Nullable;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.helpers.DefaultHandler;

import java.io.InputStream;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * The merged regions of one sheet and the value of each region's top-left cell, so that the other cells of a
 * region can be given that value while the sheet is read in order. An Excel file stores a merged region's value
 * in its top-left cell only; the rest of the region is empty.
 */
final class MergedCellFill {

    /** The most merged regions one sheet may have; a file that merges more is refused. */
    private static final int MAX_REGIONS = 200_000;

    /** A merged region, 0-based and inclusive. */
    private record Region(int firstRow, int lastRow, int firstCol, int lastCol) {}

    private record Value(@Nullable String formatted, @Nullable String raw) {}

    /** A cell to give a value: the column, the text shown and the stored number. */
    record Fill(int column, @Nullable String formatted, @Nullable String raw) {}

    private final List<Region> regions;
    /** The top-left cell of each region: row and column to the region. */
    private final Map<Long, Integer> anchors = new HashMap<>();
    private final Map<Integer, Value> values = new HashMap<>();
    private final ArrayDeque<Integer> active = new ArrayDeque<>();
    private int next;

    private MergedCellFill(List<Region> regions) {
        this.regions = regions;
        for (int i = 0; i < regions.size(); i++) {
            Region r = regions.get(i);
            anchors.put(key(r.firstRow, r.firstCol), i);
        }
    }

    private static long key(int row, int col) {
        return ((long) row << 16) | col;
    }

    /**
     * Reads the merged regions from the sheet's XML, which lists them after the rows.
     *
     * @throws ExcelReadException if the sheet merges more than {@value #MAX_REGIONS} regions
     */
    static MergedCellFill scan(InputStream sheet) throws Exception {
        var xmlReader = XMLHelper.newXMLReader();
        var collector = new Collector();
        xmlReader.setContentHandler(collector);
        xmlReader.parse(new InputSource(sheet));
        // Rows arrive in order, so the regions are opened in the order of their first row.
        collector.regions.sort(Comparator.comparingInt(Region::firstRow));
        return new MergedCellFill(List.copyOf(collector.regions));
    }

    boolean isEmpty() {
        return regions.isEmpty();
    }

    /** Called for every cell as it is read: keeps the value of a region's top-left cell. */
    void capture(int row, int col, @Nullable String formatted, @Nullable String raw) {
        Integer region = anchors.get(key(row, col));
        if (region != null) values.put(region, new Value(formatted, raw));
    }

    /**
     * The cells of {@code row} that lie in a merged region but not at its top-left corner, with the value of that
     * corner. Rows must be asked for in ascending order. A region whose top-left cell had no value gives nothing.
     */
    List<Fill> cellsOf(int row) {
        while (next < regions.size() && regions.get(next).firstRow <= row) active.add(next++);
        var fills = new ArrayList<Fill>();
        var iterator = active.iterator();
        while (iterator.hasNext()) {
            int index = iterator.next();
            Region r = regions.get(index);
            if (r.lastRow < row) {
                iterator.remove();
                values.remove(index);
                continue;
            }
            Value value = values.get(index);
            if (value == null) continue;
            for (int col = r.firstCol; col <= r.lastCol; col++) {
                if (row == r.firstRow && col == r.firstCol) continue;
                fills.add(new Fill(col, value.formatted, value.raw));
            }
        }
        return fills;
    }

    private static final class Collector extends DefaultHandler {
        private final List<Region> regions = new ArrayList<>();

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if (!"mergeCell".equals(localName) && !"mergeCell".equals(qName)) return;
            String ref = attributes.getValue("ref");
            if (ref == null) return;
            CellRangeAddress range = CellRangeAddress.valueOf(ref);
            if (range.getNumberOfCells() < 2) return;
            if (regions.size() >= MAX_REGIONS) {
                throw new ExcelReadException("The sheet merges more than " + MAX_REGIONS
                        + " regions; read it without fillMergedCells()");
            }
            regions.add(new Region(range.getFirstRow(), range.getLastRow(), range.getFirstColumn(), range.getLastColumn()));
        }
    }
}
