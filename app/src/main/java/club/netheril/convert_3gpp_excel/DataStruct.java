package club.netheril.convert_3gpp_excel;

final class ExcelCellIndex {
    private int rowValue;
    private int colValue;

    private ExcelCellIndex(int row, int col) {
        this.rowValue = row;
        this.colValue = col;
    }

    public static ExcelCellIndex of(int row, int col) {
        return new ExcelCellIndex(row, col);
    }

    public int row() {
        return rowValue;
    }

    public int col() {
        return colValue;
    }

    @Override
    public String toString() {
        return String.format("(%d, %d)", rowValue, colValue);
    }

    @Override
    public int hashCode() {
        return rowValue * 65536 + colValue;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null || getClass() != obj.getClass()) {
            return false;
        }
        ExcelCellIndex other = (ExcelCellIndex) obj;
        return row() == other.row() && col() == other.col();
    }
}

final class TableMetadata {
    // The spec name, spec version and serial number of this table in 3GPP specs,
    // e.g.: "36.101" + "h50" + "5.5A.1-1".
    private String specName;
    private String specVersion;
    private String tableSerialNumber;

    private String tableTitle;

    // Following 4 fields define a row range [begin_row, end_row) and a column
    // range [begin_column, end_column). They form a rectangle region which
    // contains all table data.
    private int beginRow;
    private int endRow;
    private int beginCol;
    private int endCol;

    public TableMetadata(String specName, String specVersion,
            String tableSerialNumber, String tableTitle,
            int beginRow, int endRow, int beginCol, int endCol) {
        this.specName = specName;
        this.specVersion = specVersion;
        this.tableSerialNumber = tableSerialNumber;
        this.tableTitle = tableTitle;
        this.beginRow = beginRow;
        this.endRow = endRow;
        this.beginCol = beginCol;
        this.endCol = endCol;
    }

    public String spec_name() {
        return specName;
    }

    public String spec_version() {
        return specVersion;
    }

    public String table_serial_number() {
        return tableSerialNumber;
    }

    public String table_title() {
        return tableTitle;
    }

    public int begin_row() {
        return beginRow;
    }

    public int end_row() {
        return endRow;
    }

    public int begin_col() {
        return beginCol;
    }

    public int end_col() {
        return endCol;
    }

}