package club.netheril.convert_3gpp_excel;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

final class CellIndex {
    public CellIndex(int row, int col) {
        this.row = Integer.valueOf(row);
        this.col = Integer.valueOf(col);
    }

    public Integer row;
    public Integer col;
}

final class SheetParserUtils {
    private static final Pattern EXCEL_CELL_NAME_PATTERN = Pattern.compile("^([A-Z]+)([0-9]+)$");

    // Parse a cell name used by Excel to the 0-based row/column index pair.
    // For example: A1 -> (0, 0), C12 -> (2, 11), AA70 -> (26, 69)
    public static CellIndex parseExcelCellName(String name) {
        Matcher m = EXCEL_CELL_NAME_PATTERN.matcher(name);
        if (!m.matches()) {
            throw new IllegalArgumentException(String.format("Unrecognizable Excel cell name '%s'", name));
        }
        return new CellIndex(Integer.valueOf(m.group(2), 10) - 1, translateColumnName(m.group(1)));
    }

    private static int translateColumnName(String name) {
        int column = 0;
        while (!name.isEmpty()) {
            column = column * 26 + (name.charAt(0) - 'A' + 1);
            name = name.substring(1);
        }
        return column - 1;
    }

}
