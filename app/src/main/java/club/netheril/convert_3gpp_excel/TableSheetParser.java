package club.netheril.convert_3gpp_excel;

import java.util.stream.IntStream;
import java.util.ArrayList;
import java.util.List;

import com.google.common.collect.ImmutableList;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.ImmutableSet;
import static com.google.common.collect.ImmutableSet.toImmutableSet;
import static com.google.common.collect.ImmutableList.toImmutableList;
import static com.google.common.collect.ImmutableMap.toImmutableMap;
import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import org.apache.commons.compress.utils.Lists;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

final class TableSheetParser {
    private static final String SHEET_NAME = "Table";

    public static TableData parse(XSSFWorkbook workbook, TableMetadata metadata) {
        XSSFSheet sheet = workbook.getSheet(SHEET_NAME);
        checkArgument(sheet != null, "Unable to find Table sheet.");
        for (int row = metadata.begin_row(); row < metadata.end_row(); row++) {
            checkArgument(
                    sheet.getRow(row) != null,
                    String.format(
                            "Invalid table sheet, row %d doesn't exist while metadata says the end row is %d",
                            row, metadata.end_row()));
            for (int column = metadata.begin_col(); column < metadata.end_col(); column++) {
                checkArgument(
                        sheet.getRow(row).getCell(column) != null,
                        String.format(
                                "Invalid table sheet, column %d at row %d doesn't exist while metadata says the end"
                                        + " column is %d",
                                column, row, metadata.end_col()));
            }
        }

        return TableData.fromRows(parseRowsFromRectagle(sheet,
                metadata.begin_row(), metadata.end_row(), metadata.begin_col(), metadata.end_col()));
    }

    // Parse a rectagle area in the sheet that consists of one or more logical rows.
    // It is exepected that:
    // 1. This rectagle is surrounded by borders.
    // 2. This rectagle is splitted into logical table rows according horizontal
    // continuous top borders that across the entire rectagle.
    private static List<TableRow> parseRowsFromRectagle(XSSFSheet sheet,
            int beginRow, int endRow, int beginColumn, int endColumn) {
        checkSheetRectagle(sheet, beginRow, endRow, beginColumn, endColumn);
        ImmutableSet<Integer> allColumns = integerSetFromRange(beginColumn, endColumn);
        checkArgument(IntStream.range(beginRow + 1, endRow)
                .anyMatch(row -> SheetParserUtils.hasTopBorder(sheet, row, allColumns)));

        int subBeginRow = beginRow;
        ArrayList<TableRow> parsedRows = Lists.newArrayList();
        for (int row = beginRow + 1; row <= endRow; row++) {
            if (SheetParserUtils.hasTopBorder(sheet, row, allColumns)) {
                parsedRows.add(parseOneRowFromRectagle(sheet, subBeginRow, row, beginColumn, endColumn));
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
    private static TableRow parseOneRowFromRectagle(XSSFSheet sheet,
            final int beginRow, final int endRow, final int beginColumn, final int endColumn) {
        checkSheetRectagle(sheet, beginRow, endRow, beginColumn, endColumn);
        for (int row = beginRow + 1; row < endRow; row++) {
            checkArgument(!SheetParserUtils.hasTopBorder(sheet, row, integerSetFromRange(beginColumn, endColumn)));
        }

        ImmutableMap<Integer, Boolean> hasSplitByColumn = IntStream.range(beginColumn, endColumn)
                .boxed()
                .collect(toImmutableMap(column -> column, column -> {
                    for (int row = beginRow + 1; row < endRow; row++) {
                        if (SheetParserUtils.hasTopBorder(sheet, row, ImmutableSet.of(column))) {
                            return true;
                        }
                    }
                    return false;
                }));

        int subBeginColumn = beginColumn;
        ArrayList<TableColumn> parsedColumns = Lists.newArrayList();
        for (int column = beginColumn + 1; column <= endColumn; column++) {
            // This is the last column or there is an internal horizontal cell boder to
            // further split.
            if (column == endColumn || hasSplitByColumn.get(column) != hasSplitByColumn.get(subBeginColumn)) {
                if (!hasSplitByColumn.get(subBeginColumn)) {
                    parsedColumns.addAll(parseLeafColumnsFromRectagle(sheet, beginRow, endRow, subBeginColumn, column));
                } else {
                    parsedColumns.add(
                            TableColumn.parent(parseRowsFromRectagle(sheet, beginRow, endRow, subBeginColumn, column)));
                }
                subBeginColumn = column;
            }
        }
        return TableRow.fromColumns(parsedColumns);
    }

    // Parse a rectagle area in the sheet that consists of only leaf logical
    // columns.
    // It is exepected that:
    // 1. This rectagle is surrounded by borders.
    // 2. There is no horizontal top borders within this rectagle.
    private static List<TableColumn> parseLeafColumnsFromRectagle(XSSFSheet sheet,
            int beginRow, int endRow, int beginColumn, int endColumn) {
        checkSheetRectagle(sheet, beginRow, endRow, beginColumn, endColumn);
        checkArgument(IntStream.range(beginRow + 1, endRow)
                .allMatch(row -> {
                    return SheetParserUtils
                            .getColumnsWithTopBorder(sheet, row, integerSetFromRange(beginColumn, endColumn))
                            .isEmpty();
                }));

        ImmutableSet<Integer> allRows = integerSetFromRange(beginRow, endRow);
        int subBeginColumn = beginColumn;
        ArrayList<TableColumn> parsedColumns = Lists.newArrayList();
        for (int column = beginColumn + 1; column <= endColumn; column++) {
            if (SheetParserUtils.getRowsWithLeftBorder(sheet, allRows, column).containsAll(allRows)) {
                parsedColumns.add(parseOneLeafColumnFromRectagle(sheet, beginRow, endRow, subBeginColumn, column));
                subBeginColumn = column;
            }
        }
        return parsedColumns;
    }

    private static TableColumn parseOneLeafColumnFromRectagle(XSSFSheet sheet,
            int beginRow, int endRow, int beginColumn, int endColumn) {
        checkSheetRectagle(sheet, beginRow, endRow, beginColumn, endColumn);

        ArrayList<ExcelCellIndex> cellIndecies = Lists.newArrayList();
        for (int row = beginRow; row < endRow; row++) {
            for (int column = beginColumn; column < endColumn; column++) {
                cellIndecies.add(ExcelCellIndex.of(row, column));
            }
        }
        ImmutableList<String> cellStrings = cellIndecies.stream()
                .map(cellIndex -> SheetParserUtils.safeGetCellString(sheet, cellIndex))
                .filter(cellString -> !cellString.isEmpty())
                .collect(toImmutableList());
        return TableColumn.leaf(cellStrings);
    }

    private static void checkSheetRectagle(XSSFSheet sheet,
            int beginRow, int endRow, int beginColumn, int endColumn) {
        checkNotNull(sheet);
        checkArgument(beginRow >= 0 && beginRow < endRow && beginColumn >= 0 && beginColumn < endColumn);

        // Checks existences of top and bottom borders.
        ImmutableSet<Integer> allColumns = integerSetFromRange(beginColumn, endColumn);
        checkArgument(
                SheetParserUtils.hasTopBorder(sheet, beginRow, allColumns),
                String.format(
                        "Invalid rectagle, there is no top border from (%d, %d) to (%d, %d)",
                        beginRow, beginColumn, beginRow, endColumn));
        checkArgument(
                SheetParserUtils.hasTopBorder(sheet, endRow, allColumns),
                String.format(
                        "Invalid rectagle, there is no top border from (%d, %d) to (%d, %d)",
                        endRow, beginColumn, endRow, endColumn));

        // Checks existences of left and right borders.
        ImmutableSet<Integer> allRows = integerSetFromRange(beginRow, endRow);
        checkArgument(
                SheetParserUtils.hasLeftBorder(sheet, allRows, beginColumn),
                String.format(
                        "Invalid rectagle, there is no left border from (%d, %d) to (%d, %d)",
                        beginRow, beginColumn, endRow, beginColumn));
        checkArgument(
                SheetParserUtils.hasLeftBorder(sheet, allRows, endColumn),
                String.format(
                        "Invalid rectagle, there is no left border from (%d, %d) to (%d, %d)",
                        beginRow, endColumn, endRow, endColumn));
    }

    private static ImmutableSet<Integer> integerSetFromRange(int begin, int end) {
        return IntStream.range(begin, end).boxed().collect(toImmutableSet());
    }
}
