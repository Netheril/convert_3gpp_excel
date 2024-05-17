package club.netheril.convert_3gpp_excel;

import static org.junit.Assert.*;

import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class MetadataSheetParserTest {

  private final String TEST_EXCEL_FILE = "table_5.3B.1.3-1.xlsx";

  private XSSFWorkbook testWorkbook;

  @Before
  public void setUp() {
    InputStream file = getClass().getClassLoader().getResourceAsStream(TEST_EXCEL_FILE);
    if (file == null) {
      throw new RuntimeException(String.format("Unable to find file '%s'", TEST_EXCEL_FILE));
    }
    try {
      testWorkbook = new XSSFWorkbook(file);
    } catch (IOException | IllegalArgumentException e) {
      throw new RuntimeException(
          String.format("Unable to read Excel from file '%s'", TEST_EXCEL_FILE), e);
    }
  }

  @After
  public void tearDown() {
    try {
      testWorkbook.close();
    } catch (IOException e) {
      throw new RuntimeException("Unable to close test workbook");
    }
  }

  @Test
  public void parse_succeed() {
    TableMetadata expected =
        new TableMetadata(
            "38.101-3",
            "17.7.0",
            "5.3B.1.3-1",
            "EN-DC configurations and bandwidth combination sets defined for intra-band"
                + " non-contiguous EN-DC",
            4,
            48,
            0,
            7);
    TableMetadata actual = MetadataSheetParser.parse(testWorkbook);
    assertEquals(expected.spec_name(), actual.spec_name());
    assertEquals(expected.spec_version(), actual.spec_version());
    assertEquals(expected.table_serial_number(), actual.table_serial_number());
    assertEquals(expected.table_title(), actual.table_title());
    assertEquals(expected.begin_row(), actual.begin_row());
    assertEquals(expected.end_row(), actual.end_row());
    assertEquals(expected.begin_col(), actual.begin_col());
    assertEquals(expected.end_col(), actual.end_col());
  }
}
