package club.netheril.convert_3gpp_excel;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import com.google.auto.value.AutoValue;
import com.google.common.base.Joiner;
import com.google.common.collect.ImmutableList;
import com.google.common.collect.Iterables;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

// This class represents a cell index in Excel sheet.
@AutoValue
abstract class ExcelCellIndex {

  private static final Pattern EXCEL_CELL_NAME_PATTERN = Pattern.compile("^([A-Z]+)([0-9]+)$");

  // Zero-based row index.
  abstract int row();

  // Zero-based column index.
  abstract int column();

  public static ExcelCellIndex of(int row, int col) {
    checkArgument(row >= 0 && col >= 0, "Row and column index must be non-negative");
    return new AutoValue_ExcelCellIndex(row, col);
  }

  public static ExcelCellIndex of(String excelCellName) {
    // Parse a cell name used by Excel to the 0-based row/column index pair.
    // For example: A1 -> (0, 0), C12 -> (2, 11), AA70 -> (26, 69)
    checkNotNull(excelCellName);
    Matcher m = EXCEL_CELL_NAME_PATTERN.matcher(excelCellName);
    checkArgument(m.matches(), String.format("Unrecognizable Excel cell name '%s'", excelCellName));
    return ExcelCellIndex.of(
        Integer.valueOf(m.group(2), 10).intValue() - 1, columnNameToIndex(m.group(1)));
  }

  @Override
  public String toString() {
    return String.format("%s%d", columnIndexToName(column() + 1), row() + 1);
  }

  private static int columnNameToIndex(String name) {
    int column = 0;
    while (!name.isEmpty()) {
      column = column * 26 + (name.charAt(0) - 'A' + 1);
      name = name.substring(1);
    }
    return column - 1;
  }

  // Please note that the column index here is 1-based.
  private static String columnIndexToName(int index) {
    if (index <= 26) {
      return String.valueOf((char) ('A' + index - 1));
    } else {
      int remainder = index % 26;
      index /= 26;
      if (remainder == 0) {
        remainder = 26;
        index -= 1;
      }
      return columnIndexToName(index) + (char) ('A' + remainder - 1);
    }
  }
}

// This class represents a rectangle region in Excel sheet.
@AutoValue
abstract class ExcelRect {
  abstract ExcelCellIndex topLeft();

  abstract ExcelCellIndex bottomRight();

  // Following 4 fields define a row range [begin_row, end_row) and a column
  // range [begin_column, end_column).
  public int beginRow() {
    return topLeft().row();
  }

  public int endRow() {
    return bottomRight().row() + 1;
  }

  public int beginColumn() {
    return topLeft().column();
  }

  public int endColumn() {
    return bottomRight().column() + 1;
  }

  public static ExcelRect of(ExcelCellIndex topLeft, ExcelCellIndex bottomRight) {
    checkNotNull(topLeft);
    checkNotNull(bottomRight);
    checkArgument(
        topLeft.row() <= bottomRight.row() && topLeft.column() <= bottomRight.column(),
        String.format(
            "Invalid Excel rect: top left cell %s is not above or to the left of bottom right cell"
                + " %s",
            topLeft, bottomRight));
    return new AutoValue_ExcelRect(topLeft, bottomRight);
  }

  public static ExcelRect of(String topLeftName, String bottomRightName) {
    return ExcelRect.of(ExcelCellIndex.of(topLeftName), ExcelCellIndex.of(bottomRightName));
  }

  public static ExcelRect of(int beginRow, int endRow, int beginColumn, int endColumn) {
    return ExcelRect.of(
        ExcelCellIndex.of(beginRow, beginColumn), ExcelCellIndex.of(endRow - 1, endColumn - 1));
  }

  @Override
  public String toString() {
    return String.format("<%s, %s>", topLeft(), bottomRight());
  }
}

@AutoValue
abstract class TableMetadata {
  abstract String specName(); // The spec name, e.g., "36.101"

  abstract String specVersion(); // The spec version, e.g., "h50"

  abstract String tableSerialNumber(); // The table serial number, e.g., "5.5A.1-1"

  abstract String tableTitle(); // The table title.

  abstract ExcelRect tableDataRect(); // The rect in Excel file which contains all table data.

  static Builder builder() {
    return new AutoValue_TableMetadata.Builder();
  }

  @AutoValue.Builder
  abstract static class Builder {
    abstract Builder setSpecName(String specName);

    abstract Builder setSpecVersion(String specVersion);

    abstract Builder setTableSerialNumber(String tableSerialNumber);

    abstract Builder setTableTitle(String tableTitle);

    abstract Builder setTableDataRect(ExcelRect tableDataRect);

    abstract TableMetadata build();
  }
}

final class TableRow {
  private ImmutableList<TableColumn> columnList;

  private TableRow(ImmutableList<TableColumn> columns) {
    columnList = columns;
  }

  public static TableRow of(List<TableColumn> columns) {
    checkArgument(columns != null && !columns.isEmpty());
    return new TableRow(ImmutableList.copyOf(columns));
  }

  public static TableRow of(TableColumn... columns) {
    return TableRow.of(ImmutableList.copyOf(columns));
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
    if (this == obj) return true;
    if (obj == null || getClass() != obj.getClass()) return false;
    TableRow other = (TableRow) obj;
    if (columnList == null) {
      return other.columnList == null;
    } else {
      return Iterables.elementsEqual(columnList, other.columnList);
    }
  }

  @Override
  public String toString() {
    return "TableRow ["
        + Joiner.on(", ").join(columnList.stream().map(c -> c.toString()).toArray())
        + "]";
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
    return cells == null || cells.isEmpty()
        ? TableColumn.empty()
        : new TableColumn("leaf", ImmutableList.copyOf(cells), null);
  }

  public static TableColumn leaf(String... cells) {
    return TableColumn.leaf(ImmutableList.copyOf(cells));
  }

  public static TableColumn parent(List<TableRow> childRows) {
    return childRows == null || childRows.isEmpty()
        ? TableColumn.empty()
        : new TableColumn("parent", null, ImmutableList.copyOf(childRows));
  }

  public static TableColumn parent(TableRow... childRows) {
    return TableColumn.parent(ImmutableList.copyOf(childRows));
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
    if (this == obj) return true;
    if (obj == null || getClass() != obj.getClass()) return false;
    TableColumn other = (TableColumn) obj;
    if (!type.equals(other.type)) return false;
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
          + Joiner.on(", ").join(childRowList.stream().map(child -> child.toString()).toArray())
          + "]";
    }
    throw new AssertionError(String.format("Unknown column type %s", type));
  }
}

final class TableData {
  private final ImmutableList<TableRow> rowList;

  private TableData(ImmutableList<TableRow> rows) {
    rowList = rows;
  }

  public static TableData of(List<TableRow> rows) {
    return new TableData(ImmutableList.copyOf(rows));
  }

  public static TableData of(TableRow... rows) {
    return TableData.of(ImmutableList.copyOf(rows));
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
    if (this == obj) return true;
    if (obj == null || getClass() != obj.getClass()) return false;
    TableData other = (TableData) obj;
    if (rowList == null) {
      return other.rowList == null;
    } else {
      return rowList.equals(other.rowList);
    }
  }
}
