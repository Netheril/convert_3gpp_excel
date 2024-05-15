package club.netheril.convert_3gpp_excel;

import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;

import com.google.common.base.Joiner;
import com.google.common.base.Optional;
import com.google.common.collect.ImmutableList;

final class SheetParserUtils {

    private static final int MAX_COLUMN_IDX = 20;
    private static final Pattern EXCEL_CELL_NAME_PATTERN = Pattern.compile("^([A-Z]+)([0-9]+)$");

    // Parse a cell name used by Excel to the 0-based row/column index pair.
    // For example: A1 -> (0, 0), C12 -> (2, 11), AA70 -> (26, 69)
    public static ExcelCellIndex parseExcelCellName(String name) {
        Matcher m = EXCEL_CELL_NAME_PATTERN.matcher(name);
        if (!m.matches()) {
            throw new IllegalArgumentException(String.format("Unrecognizable Excel cell name '%s'", name));
        }
        return new ExcelCellIndex(Integer.valueOf(m.group(2), 10) - 1, translateColumnName(m.group(1)));
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
                        // Superscription or subscription, treat it as a separator.
                        if (font.getTypeOffset() != Font.SS_SUPER && font.getTypeOffset() != Font.SS_SUB) {
                            // This is not a superscription or subscription, break for unexpected format.
                            throw new IllegalArgumentException(
                                    String.format("Unsupported rich text format at cell %s", cellIndex.toString()));
                        }
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
        if (sheet == null
                // We have a "+1" here as we use "the row after the last row" as the end mark.
                || cellIndex.row() < 0 || cellIndex.row() > sheet.getLastRowNum() + 1
                // Apache POI doesn't support get last column number, use a reasonable constant
                // instead.
                || cellIndex.col() < 0 || cellIndex.col() > MAX_COLUMN_IDX) {
            throw new IllegalArgumentException(
                    String.format("Invalid cell index %s", cellIndex.toString()));
        }
    }

    private static Optional<XSSFCell> safeGetCell(XSSFSheet sheet, ExcelCellIndex cellIndex) {
        checkSheetArgument(sheet, cellIndex);
        Optional<XSSFRow> row = Optional.fromNullable(sheet.getRow(cellIndex.row()));
        return row.isPresent() ? Optional.fromNullable(row.get().getCell(cellIndex.col())) : Optional.absent();
    }
}
