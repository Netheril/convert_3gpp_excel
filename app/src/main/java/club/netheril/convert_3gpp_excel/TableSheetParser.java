package club.netheril.convert_3gpp_excel;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;
import static com.google.common.collect.ImmutableList.toImmutableList;
import static com.google.common.collect.ImmutableMap.toImmutableMap;
import static com.google.common.collect.ImmutableSet.toImmutableSet;

import com.google.common.collect.ImmutableList;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.ImmutableSet;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

final class TableSheetParser {
  private static final String SHEET_NAME = "Table";

  public static TableData parse(XSSFWorkbook workbook, TableMetadata metadata) {
    XSSFSheet sheet = workbook.getSheet(SHEET_NAME);
    checkArgument(sheet != null, "Unable to find Table sheet.");
    ExcelRect tableDataRect = metadata.tableDataRect();
    for (int row = tableDataRect.beginRow(); row < tableDataRect.endRow(); row++) {
      for (int column = tableDataRect.beginColumn(); column < tableDataRect.endColumn(); column++) {
        checkArgument(
            sheet.getRow(row) != null && sheet.getRow(row).getCell(column) != null,
            String.format(
                "Invalid table sheet, cell %s in table data rect %s doesn't exist",
                ExcelCellIndex.of(row, column), tableDataRect));
      }
    }

    return TableData.of(parseRowsFromRectagle(sheet, metadata.tableDataRect(), false));
  }

  // Parse a rectagle area in the sheet that consists of one or more logical rows.
  // It is exepected that:
  // 1. This rectagle is surrounded by borders.
  // 2. This rectagle is splitted into logical table rows according horizontal
  // continuous top borders that across the entire rectagle.
  private static List<TableRow> parseRowsFromRectagle(
      XSSFSheet sheet, ExcelRect rect, boolean strictCheck) {
    checkSheetRectagle(sheet, rect);
    ImmutableSet<Integer> allColumns = integerSetFromRange(rect.beginColumn(), rect.endColumn());
    if (strictCheck) {
      checkArgument(
          IntStream.range(rect.beginRow() + 1, rect.endRow())
              .anyMatch(row -> SheetParserUtils.hasTopBorder(sheet, row, allColumns)));
    }

    int subBeginRow = rect.beginRow();
    ArrayList<TableRow> parsedRows = Lists.newArrayList();
    for (int row = rect.beginRow() + 1; row <= rect.endRow(); row++) {
      if (SheetParserUtils.hasTopBorder(sheet, row, allColumns)) {
        parsedRows.add(
            parseOneRowFromRectagle(
                sheet, ExcelRect.of(subBeginRow, row, rect.beginColumn(), rect.endColumn())));
        subBeginRow = row;
      }
    }
    return parsedRows;
  }

  // Parse a rectagle area in the sheet that consists of only one logical row.
  // It is exepected that:
  // 1. This rectagle is surrounded by borders.
  // 2. There is no horizontal continuous top borders that across the entire
  // rectagle.
  private static TableRow parseOneRowFromRectagle(XSSFSheet sheet, ExcelRect rect) {
    checkSheetRectagle(sheet, rect);
    for (int row = rect.beginRow() + 1; row < rect.endRow(); row++) {
      checkArgument(
          !SheetParserUtils.hasTopBorder(
              sheet, row, integerSetFromRange(rect.beginColumn(), rect.endColumn())));
    }

    ImmutableMap<Integer, Boolean> hasSplitByColumn =
        IntStream.range(rect.beginColumn(), rect.endColumn())
            .boxed()
            .collect(
                toImmutableMap(
                    column -> column,
                    column -> {
                      for (int row = rect.beginRow() + 1; row < rect.endRow(); row++) {
                        if (SheetParserUtils.hasTopBorder(sheet, row, ImmutableSet.of(column))) {
                          return true;
                        }
                      }
                      return false;
                    }));

    int subBeginColumn = rect.beginColumn();
    ArrayList<TableColumn> parsedColumns = Lists.newArrayList();
    for (int column = rect.beginColumn() + 1; column <= rect.endColumn(); column++) {
      // This is the last column or there is an internal horizontal cell boder to
      // further split.
      if (column == rect.endColumn()
          || hasSplitByColumn.get(column) != hasSplitByColumn.get(subBeginColumn)) {
        if (!hasSplitByColumn.get(subBeginColumn)) {
          parsedColumns.addAll(
              parseLeafColumnsFromRectagle(
                  sheet, ExcelRect.of(rect.beginRow(), rect.endRow(), subBeginColumn, column)));
        } else {
          parsedColumns.add(
              TableColumn.parent(
                  parseRowsFromRectagle(
                      sheet,
                      ExcelRect.of(rect.beginRow(), rect.endRow(), subBeginColumn, column),
                      true)));
        }
        subBeginColumn = column;
      }
    }
    return TableRow.of(parsedColumns);
  }

  // Parse a rectagle area in the sheet that consists of only leaf logical
  // columns.
  // It is exepected that:
  // 1. This rectagle is surrounded by borders.
  // 2. There is no horizontal top borders within this rectagle.
  private static List<TableColumn> parseLeafColumnsFromRectagle(XSSFSheet sheet, ExcelRect rect) {
    checkSheetRectagle(sheet, rect);
    checkArgument(
        IntStream.range(rect.beginRow() + 1, rect.endRow())
            .allMatch(
                row -> {
                  return SheetParserUtils.getColumnsWithTopBorder(
                          sheet, row, integerSetFromRange(rect.beginColumn(), rect.endColumn()))
                      .isEmpty();
                }));

    ImmutableSet<Integer> allRows = integerSetFromRange(rect.beginRow(), rect.endRow());
    int subBeginColumn = rect.beginColumn();
    ArrayList<TableColumn> parsedColumns = Lists.newArrayList();
    for (int column = rect.beginColumn() + 1; column <= rect.endColumn(); column++) {
      if (SheetParserUtils.getRowsWithLeftBorder(sheet, allRows, column).containsAll(allRows)) {
        parsedColumns.add(
            parseOneLeafColumnFromRectagle(
                sheet, ExcelRect.of(rect.beginRow(), rect.endRow(), subBeginColumn, column)));
        subBeginColumn = column;
      }
    }
    return parsedColumns;
  }

  private static TableColumn parseOneLeafColumnFromRectagle(XSSFSheet sheet, ExcelRect rect) {
    checkSheetRectagle(sheet, rect);

    ArrayList<ExcelCellIndex> cellIndecies = Lists.newArrayList();
    for (int row = rect.beginRow(); row < rect.endRow(); row++) {
      for (int column = rect.beginColumn(); column < rect.endColumn(); column++) {
        cellIndecies.add(ExcelCellIndex.of(row, column));
      }
    }
    ImmutableList<String> cellStrings =
        cellIndecies.stream()
            .map(cellIndex -> SheetParserUtils.safeGetCellString(sheet, cellIndex))
            .filter(cellString -> !cellString.isEmpty())
            .collect(toImmutableList());
    return TableColumn.leaf(cellStrings);
  }

  private static void checkSheetRectagle(XSSFSheet sheet, ExcelRect rect) {
    checkNotNull(sheet);
    checkNotNull(rect);

    // Checks existences of top and bottom borders.
    ImmutableSet<Integer> allColumns = integerSetFromRange(rect.beginColumn(), rect.endColumn());
    checkArgument(
        SheetParserUtils.hasTopBorder(sheet, rect.beginRow(), allColumns),
        String.format("Invalid rectagle %s, it has no top border", rect));

    checkArgument(
        SheetParserUtils.hasTopBorder(sheet, rect.endRow(), allColumns),
        String.format("Invalid rectagle %s, it has no bottom border", rect));

    // Checks existences of left and right borders.
    ImmutableSet<Integer> allRows = integerSetFromRange(rect.beginRow(), rect.endRow());
    checkArgument(
        SheetParserUtils.hasLeftBorder(sheet, allRows, rect.beginColumn()),
        String.format("Invalid rectagle %s, it has no left border", rect));
    checkArgument(
        SheetParserUtils.hasLeftBorder(sheet, allRows, rect.endColumn()),
        String.format("Invalid rectagle %s, it has no right border", rect));
  }

  private static ImmutableSet<Integer> integerSetFromRange(int begin, int end) {
    return IntStream.range(begin, end).boxed().collect(toImmutableSet());
  }
}
