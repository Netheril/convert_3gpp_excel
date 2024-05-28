package club.netheril.convert_3gpp_excel;

import static com.google.common.base.Preconditions.checkNotNull;
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
    checkNotNull(file, "Unable to find file '%s'", TEST_EXCEL_FILE);
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
        TableMetadata.builder()
            .setSpecName("38.101-3")
            .setSpecVersion("17.7.0")
            .setTableSerialNumber("5.3B.1.3-1")
            .setTableTitle(
                "EN-DC configurations and bandwidth combination sets defined for intra-band"
                    + " non-contiguous EN-DC")
            .setTableDataRect(ExcelRect.of("A5", "G48"))
            .build();
    TableMetadata actual = MetadataSheetParser.parse(testWorkbook);
    assertEquals(expected, actual);
  }
}
