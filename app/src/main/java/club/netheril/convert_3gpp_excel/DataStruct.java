package club.netheril.convert_3gpp_excel;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import com.google.auto.value.AutoValue;
import com.google.auto.value.extension.toprettystring.ToPrettyString;
import com.google.common.base.Joiner;
import com.google.common.collect.ImmutableList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.annotation.Nullable;

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

@AutoValue
abstract class TableRow {
  abstract ImmutableList<TableColumn> columns();

  public static TableRow of(List<TableColumn> columns) {
    checkArgument(columns != null && !columns.isEmpty());
    return new AutoValue_TableRow(ImmutableList.copyOf(columns));
  }

  public static TableRow of(TableColumn... columns) {
    return TableRow.of(ImmutableList.copyOf(columns));
  }

  @ToPrettyString
  @Override
  public abstract String toString();
}

@AutoValue
abstract class TableColumn {
  enum Type {
    EMPTY,
    LEAF,
    PARENT
  }

  abstract Type type();

  @Nullable
  abstract ImmutableList<String> cells();

  @Nullable
  abstract ImmutableList<TableRow> childRows();

  private static final TableColumn EMPTY = new AutoValue_TableColumn(Type.EMPTY, null, null);

  public static TableColumn empty() {
    return EMPTY;
  }

  public static TableColumn leaf(List<String> cells) {
    return cells == null || cells.isEmpty()
        ? TableColumn.empty()
        : new AutoValue_TableColumn(Type.LEAF, ImmutableList.copyOf(cells), null);
  }

  public static TableColumn leaf(String... cells) {
    return TableColumn.leaf(ImmutableList.copyOf(cells));
  }

  public static TableColumn parent(List<TableRow> childRows) {
    return childRows == null || childRows.isEmpty()
        ? TableColumn.empty()
        : new AutoValue_TableColumn(Type.PARENT, null, ImmutableList.copyOf(childRows));
  }

  public static TableColumn parent(TableRow... childRows) {
    return TableColumn.parent(ImmutableList.copyOf(childRows));
  }

  @Override
  public String toString() {
    switch (type()) {
      case EMPTY:
        return "EmptyColumn";
      case LEAF:
        return "LeafColumn [" + Joiner.on(", ").join(cells()) + "]";
      case PARENT:
        return "ParentColumn ["
            + Joiner.on(", ").join(childRows().stream().map(TableRow::toString).iterator())
            + "]";
      default:
        throw new IllegalStateException("Unknown TableColumn type: " + type());
    }
  }
}

@AutoValue
abstract class TableData {
  abstract ImmutableList<TableRow> rows();

  public static TableData of(List<TableRow> rows) {
    checkArgument(rows != null && !rows.isEmpty());
    return new AutoValue_TableData(ImmutableList.copyOf(rows));
  }

  public static TableData of(TableRow... rows) {
    return TableData.of(ImmutableList.copyOf(rows));
  }

  @ToPrettyString
  @Override
  public abstract String toString();
}
