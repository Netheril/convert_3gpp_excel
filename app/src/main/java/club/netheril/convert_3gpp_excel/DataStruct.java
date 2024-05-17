package club.netheril.convert_3gpp_excel;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import java.util.List;

import com.google.common.base.Joiner;
import com.google.common.collect.ImmutableList;
import com.google.common.collect.Iterables;

final class ExcelCellIndex {
    private int rowValue;
    private int colValue;

    private ExcelCellIndex(int row, int col) {
        checkArgument(row >= 0 && col >= 0);
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
    private final String specName;
    private final String specVersion;
    private final String tableSerialNumber;

    private final String tableTitle;

    // Following 4 fields define a row range [begin_row, end_row) and a column
    // range [begin_column, end_column). They form a rectangle region which
    // contains all table data.
    private final int beginRow;
    private final int endRow;
    private final int beginCol;
    private final int endCol;

    public TableMetadata(String specName, String specVersion,
            String tableSerialNumber, String tableTitle,
            int beginRow, int endRow, int beginCol, int endCol) {
        checkArgument(beginRow >= 0 && beginRow < endRow && beginCol >= 0 && beginCol < endCol);
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

final class TableRow {
    private ImmutableList<TableColumn> columnList;

    private TableRow(ImmutableList<TableColumn> columns) {
        columnList = columns;
    }

    public static TableRow fromColumns(List<TableColumn> columns) {
        checkArgument(columns != null && !columns.isEmpty());
        return new TableRow(ImmutableList.copyOf(columns));
    }

    public ImmutableList<TableColumn> columns() {
        return columnList;
    }

    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result + ((columnList == null) ? 0 : columnList.hashCode());
        return result;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj)
            return true;
        if (obj == null || getClass() != obj.getClass())
            return false;
        TableRow other = (TableRow) obj;
        if (columnList == null) {
            return other.columnList == null;
        } else {
            return Iterables.elementsEqual(columnList, other.columnList);
        }
    }

    @Override
    public String toString() {
        return "TableRow [" + Joiner.on(", ").join(columnList.stream().map(c -> c.toString()).toArray()) + "]";
    }
}

final class TableColumn {
    private static final TableColumn EMPTY = new TableColumn("empty", null, null);

    private final String type;
    private final ImmutableList<String> cellList;
    private final ImmutableList<TableRow> childRowList;

    private TableColumn(String type, ImmutableList<String> cells, ImmutableList<TableRow> childRows) {
        this.type = type;
        cellList = cells;
        childRowList = childRows;
    }

    public static TableColumn empty() {
        return EMPTY;
    }

    public static TableColumn leaf(List<String> cells) {
        if (cells != null && !cells.isEmpty()) {
            return new TableColumn("leaf", ImmutableList.copyOf(cells), null);
        } else {
            return TableColumn.empty();
        }
    }

    public static TableColumn parent(List<TableRow> childRows) {
        if (childRows != null && !childRows.isEmpty()) {
            return new TableColumn("parent", null, ImmutableList.copyOf(childRows));
        } else {
            return TableColumn.empty();
        }
    }

    public ImmutableList<String> cells() {
        checkNotNull(cellList, String.format("No cell values available for this %s column!", type));
        return cellList;
    }

    public ImmutableList<TableRow> child_rows() {
        checkNotNull(childRowList, String.format("No child rows available for this %s column!", type));
        return childRowList;
    }

    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result + ((type == null) ? 0 : type.hashCode());
        result = prime * result + ((cellList == null) ? 0 : cellList.hashCode());
        result = prime * result + ((childRowList == null) ? 0 : childRowList.hashCode());
        return result;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj)
            return true;
        if (obj == null || getClass() != obj.getClass())
            return false;
        TableColumn other = (TableColumn) obj;
        if (!type.equals(other.type))
            return false;
        if (type == "empty") {
            return true;
        } else if (type == "leaf") {
            return Iterables.elementsEqual(cellList, other.cellList);
        } else if (type == "parent") {
            return Iterables.elementsEqual(childRowList, other.childRowList);
        }
        throw new AssertionError(String.format("Unknown column type %s", type));
    }

    @Override
    public String toString() {
        if (type == "empty") {
            return "TableEmptyColumn";
        } else if (type == "leaf") {
            return "TableLeafColumn [" + Joiner.on(", ").join(cellList) + "]";
        } else if (type == "parent") {
            return "TableParentColumn ["
                    + Joiner.on(", ").join(childRowList.stream().map(child -> child.toString()).toArray()) + "]";
        }
        throw new AssertionError(String.format("Unknown column type %s", type));
    }
}

final class TableData {
    private final ImmutableList<TableRow> rowList;

    private TableData(ImmutableList<TableRow> rows) {
        rowList = rows;
    }

    public static TableData fromRows(List<TableRow> rows) {
        return new TableData(ImmutableList.copyOf(rows));
    }

    public ImmutableList<TableRow> rows() {
        return rowList;
    }

    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result + ((rowList == null) ? 0 : rowList.hashCode());
        return result;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj)
            return true;
        if (obj == null || getClass() != obj.getClass())
            return false;
        TableData other = (TableData) obj;
        if (rowList == null) {
            return other.rowList == null;
        } else {
            return rowList.equals(other.rowList);
        }
    }
}