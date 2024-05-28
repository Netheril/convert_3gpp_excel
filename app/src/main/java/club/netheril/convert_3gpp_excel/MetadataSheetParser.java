package club.netheril.convert_3gpp_excel;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

final class MetadataSheetParser {
  private static final String SHEET_NAME = "Metadata";

  private static final String KEY_SPEC_NAME = "Spec";
  private static final String KEY_SPEC_VERSION = "Version";
  private static final String KEY_SERIAL_NUM = "Number";
  private static final String KEY_TITLE = "Title";
  private static final String KEY_TOP_LEFT = "Top left";
  private static final String KEY_TOP_BOTTOM_RIGHT = "Bottom right";

  public static TableMetadata parse(XSSFWorkbook workbook) {
    XSSFSheet sheet = workbook.getSheet(SHEET_NAME);
    checkNotNull(sheet, "Unable to find metadata sheet in the Excel file");

    TableMetadata.Builder builder = TableMetadata.builder();
    ExcelCellIndex topLeft = null;
    ExcelCellIndex bottomRight = null;
    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
      String key = SheetParserUtils.safeGetCellString(sheet, ExcelCellIndex.of(i, 0));
      String value = SheetParserUtils.safeGetCellString(sheet, ExcelCellIndex.of(i, 1));
      if (key.isEmpty() && value.isEmpty()) {
        continue;
      }
      if (key.equals(KEY_SPEC_NAME)) {
        builder.setSpecName(value);
      } else if (key.equals(KEY_SPEC_VERSION)) {
        builder.setSpecVersion(value);
      } else if (key.equals(KEY_SERIAL_NUM)) {
        builder.setTableSerialNumber(value);
      } else if (key.equals(KEY_TITLE)) {
        builder.setTableTitle(value);
      } else if (key.equals(KEY_TOP_LEFT)) {
        topLeft = ExcelCellIndex.of(value);
      } else if (key.equals(KEY_TOP_BOTTOM_RIGHT)) {
        bottomRight = ExcelCellIndex.of(value);
      } else {
        throw new IllegalArgumentException(
            String.format("Unrecognizable key column '%s' in metadata sheet", key));
      }
    }
    checkArgument(
        topLeft != null && bottomRight != null,
        "Missing top left and/or bottom right cell of the table data rect");

    return builder.setTableDataRect(ExcelRect.of(topLeft, bottomRight)).build();
  }
}
