package club.netheril.convert_3gpp_excel;

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

    String specName = null;
    String specVersion = null;
    String tableSerialNumber = null;
    String tableTitle = null;
    int beginRow = -1;
    int endRow = -1;
    int beginCol = -1;
    int endCol = -1;

    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
      String key = SheetParserUtils.safeGetCellString(sheet, ExcelCellIndex.of(i, 0));
      String value = SheetParserUtils.safeGetCellString(sheet, ExcelCellIndex.of(i, 1));
      if (key.isEmpty() && value.isEmpty()) {
        continue;
      }
      if (key.equals(KEY_SPEC_NAME)) {
        specName = value;
      } else if (key.equals(KEY_SPEC_VERSION)) {
        specVersion = value;
      } else if (key.equals(KEY_SERIAL_NUM)) {
        tableSerialNumber = value;
      } else if (key.equals(KEY_TITLE)) {
        tableTitle = value;
      } else if (key.equals(KEY_TOP_LEFT)) {
        ExcelCellIndex topLeft = SheetParserUtils.parseExcelCellName(value);
        beginRow = topLeft.row();
        beginCol = topLeft.col();
      } else if (key.equals(KEY_TOP_BOTTOM_RIGHT)) {
        ExcelCellIndex bottomRight = SheetParserUtils.parseExcelCellName(value);
        endRow = bottomRight.row() + 1;
        endCol = bottomRight.col() + 1;
      } else {
        throw new IllegalArgumentException(
            String.format("Unrecognizable key column '%s' in metadata sheet", key));
      }
    }

    return new TableMetadata(
        specName, specVersion, tableSerialNumber, tableTitle, beginRow, endRow, beginCol, endCol);
  }
}
