package club.netheril.convert_3gpp_excel;

import static com.google.common.base.Preconditions.*;
import static com.google.common.collect.ImmutableSet.toImmutableSet;

import com.google.common.base.Joiner;
import com.google.common.base.Optional;
import com.google.common.collect.ImmutableList;
import com.google.common.collect.ImmutableSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

final class SheetParserUtils {

  private static final int MAX_COLUMN_IDX = 1000;
  private static final Pattern EXCEL_CELL_NAME_PATTERN = Pattern.compile("^([A-Z]+)([0-9]+)$");

  // Parse a cell name used by Excel to the 0-based row/column index pair.
  // For example: A1 -> (0, 0), C12 -> (2, 11), AA70 -> (26, 69)
  public static ExcelCellIndex parseExcelCellName(String name) {
    Matcher m = EXCEL_CELL_NAME_PATTERN.matcher(name);
    checkArgument(m.matches(), String.format("Unrecognizable Excel cell name '%s'", name));
    return ExcelCellIndex.of(
        Integer.valueOf(m.group(2), 10).intValue() - 1, translateColumnName(m.group(1)));
  }

  private static int translateColumnName(String name) {
    int column = 0;
    while (!name.isEmpty()) {
      column = column * 26 + (name.charAt(0) - 'A' + 1);
      name = name.substring(1);
    }
    return column - 1;
  }

  public static String safeGetCellString(XSSFSheet sheet, ExcelCellIndex cellIndex) {
    Optional<XSSFCell> cell = safeGetCell(sheet, cellIndex);
    if (!cell.isPresent()) {
      return "";
    }

    switch (cell.get().getCellType()) {
      case CellType.BLANK:
        return "";
      case CellType.BOOLEAN:
        return cell.get().getBooleanCellValue() ? "true" : "false";
      case CellType.NUMERIC:
        double value = cell.get().getNumericCellValue();
        return value % 1 == 0 ? String.format("%d", (long) value) : String.valueOf(value);
      case CellType.STRING:
        XSSFRichTextString richTextString = cell.get().getRichStringCellValue();
        // Regular text cell, no special format applied on this cell.
        if (richTextString.numFormattingRuns() == 0 || richTextString.numFormattingRuns() == 1) {
          return richTextString.toString().trim();
        }
        // Rich format text, we handle each rich format text run and join them together.
        StringBuilder runBuilder = new StringBuilder();
        ImmutableList.Builder<String> formattingRuns = new ImmutableList.Builder<>();
        for (int i = 0; i < richTextString.numFormattingRuns(); i++) {
          XSSFFont font = richTextString.getFontOfFormattingRun(i);
          if (font == null || font.getTypeOffset() == Font.SS_NONE) {
            // Regular text, merged together if they are not separated by superscriptions or
            // subscriptions.
            runBuilder.append(
                richTextString
                    .toString()
                    .substring(
                        richTextString.getIndexOfFormattingRun(i),
                        richTextString.getIndexOfFormattingRun(i)
                            + richTextString.getLengthOfFormattingRun(i))
                    .trim());
          } else {
            // This must be a superscription or a subscription, treat it as a separator.
            checkArgument(
                font.getTypeOffset() == Font.SS_SUPER || font.getTypeOffset() == Font.SS_SUB,
                String.format("Unsupported rich text format at cell %s", cellIndex.toString()));
            if (runBuilder.length() != 0) {
              formattingRuns.add(runBuilder.toString());
              runBuilder = new StringBuilder();
            }
          }
        }
        if (runBuilder.length() != 0) {
          formattingRuns.add(runBuilder.toString());
        }
        return Joiner.on(",").join(formattingRuns.build());

      default:
        break;
    }
    throw new IllegalArgumentException(
        String.format(
            "Unsupported cell type %s at %s with text %s.",
            cell.get().getCellType(), cellIndex.toString(), cell.get().toString()));
  }

  private static void checkSheetArgument(XSSFSheet sheet, ExcelCellIndex cellIndex) {
    checkArgument(
        sheet != null
            // We have a "+1" here as we use "the row after the last row" as the end mark.
            && cellIndex.row() >= 0
            && cellIndex.row() <= sheet.getLastRowNum() + 1
            // Apache POI doesn't support get last column number, use a reasonable constant
            // instead.
            && cellIndex.col() >= 0
            && cellIndex.col() <= MAX_COLUMN_IDX,
        String.format("Invalid cell index %s", cellIndex.toString()));
  }

  private static Optional<XSSFCell> safeGetCell(XSSFSheet sheet, ExcelCellIndex cellIndex) {
    checkSheetArgument(sheet, cellIndex);
    Optional<XSSFRow> row = Optional.fromNullable(sheet.getRow(cellIndex.row()));
    return row.isPresent()
        ? Optional.fromNullable(row.get().getCell(cellIndex.col()))
        : Optional.absent();
  }

  // At given row and within given columns, find columns which have top border.
  public static ImmutableSet<Integer> getColumnsWithTopBorder(
      XSSFSheet sheet, int row, ImmutableSet<Integer> allColumns) {
    checkArgument(!allColumns.isEmpty());
    return allColumns.stream()
        .filter(
            column -> {
              checkSheetArgument(sheet, ExcelCellIndex.of(row, column.intValue()));
              // By definition, all cells in the first row has a top border.
              if (row == 0) {
                return true;
              }
              Optional<XSSFCell> cell =
                  safeGetCell(sheet, ExcelCellIndex.of(row, column.intValue()));
              if (cell.isPresent() && cellHasTopBorder(cell.get())) {
                return true;
              }
              Optional<XSSFCell> prevRowCell =
                  safeGetCell(sheet, ExcelCellIndex.of(row - 1, column.intValue()));
              return prevRowCell.isPresent() && cellHasBottomBorder(prevRowCell.get());
            })
        .collect(toImmutableSet());
  }

  // For all cells in the given sheet at the given row and the given columns have
  // a "top border".
  public static boolean hasTopBorder(XSSFSheet sheet, int row, ImmutableSet<Integer> columns) {
    return getColumnsWithTopBorder(sheet, row, columns).size() == columns.size();
  }

  // At given column and within given rows, find rows which have left border.
  public static ImmutableSet<Integer> getRowsWithLeftBorder(
      XSSFSheet sheet, ImmutableSet<Integer> allRows, int column) {
    checkArgument(!allRows.isEmpty());
    return allRows.stream()
        .filter(
            row -> {
              checkSheetArgument(sheet, ExcelCellIndex.of(row.intValue(), column));
              // By definition, all cells in the first column has a left border.
              if (column == 0) {
                return true;
              }
              Optional<XSSFCell> cell =
                  safeGetCell(sheet, ExcelCellIndex.of(row.intValue(), column));
              if (cell.isPresent() && cellHasLeftBorder(cell.get())) {
                return true;
              }
              Optional<XSSFCell> prevColumnCell =
                  safeGetCell(sheet, ExcelCellIndex.of(row.intValue(), column - 1));
              return prevColumnCell.isPresent() && cellHasRightBorder(prevColumnCell.get());
            })
        .collect(toImmutableSet());
  }

  // Whether all cells in sheet at the given rows and column have a "left border".
  public static boolean hasLeftBorder(XSSFSheet sheet, ImmutableSet<Integer> rows, int column) {
    return getRowsWithLeftBorder(sheet, rows, column).size() == rows.size();
  }

  private static boolean cellHasTopBorder(XSSFCell cell) {
    XSSFCellStyle style = checkNotNull((XSSFCellStyle) cell.getCellStyle());
    return style.getBorderTop() != BorderStyle.NONE
        && isBorderColorVisible(style.getTopBorderXSSFColor());
  }

  private static boolean cellHasBottomBorder(XSSFCell cell) {
    XSSFCellStyle style = checkNotNull((XSSFCellStyle) cell.getCellStyle());
    return style.getBorderBottom() != BorderStyle.NONE
        && isBorderColorVisible(style.getBottomBorderXSSFColor());
  }

  private static boolean cellHasLeftBorder(XSSFCell cell) {
    XSSFCellStyle style = checkNotNull((XSSFCellStyle) cell.getCellStyle());
    return style.getBorderLeft() != BorderStyle.NONE
        && isBorderColorVisible(style.getLeftBorderXSSFColor());
  }

  private static boolean cellHasRightBorder(XSSFCell cell) {
    XSSFCellStyle style = checkNotNull((XSSFCellStyle) cell.getCellStyle());
    return style.getBorderRight() != BorderStyle.NONE
        && isBorderColorVisible(style.getRightBorderXSSFColor());
  }

  // Whether cell border color is visible on a white background.
  private static boolean isBorderColorVisible(XSSFColor color) {
    if (color.getRGB() == null) {
      return true;
    }
    for (byte b : color.getRGB()) {
      // Any colors "darker" than [250, 250, 250] are considered visible.
      if ((b & 0xFF) < 250) {
        return true;
      }
    }
    return false;
  }
}
