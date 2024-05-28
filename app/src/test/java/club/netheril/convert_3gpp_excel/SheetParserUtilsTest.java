package club.netheril.convert_3gpp_excel;

import static com.google.common.base.Preconditions.checkNotNull;
import static org.junit.Assert.*;

import com.google.common.collect.ImmutableMap;
import com.google.common.collect.ImmutableSet;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class SheetParserUtilsTest {

  private final String TEST_EXCEL_FILE = "table_border_test_data.xlsx";

  private XSSFWorkbook testWorkbook;
  private XSSFSheet testSheet;

  @Before
  public void setUp() {
    InputStream file = getClass().getClassLoader().getResourceAsStream(TEST_EXCEL_FILE);
    checkNotNull(file, String.format("Unable to find file '%s'", TEST_EXCEL_FILE));
    try {
      testWorkbook = new XSSFWorkbook(file);
    } catch (IOException | IllegalArgumentException e) {
      throw new IllegalArgumentException(
          String.format("Unable to read Excel from file '%s'", TEST_EXCEL_FILE), e);
    }
    testSheet = testWorkbook.getSheetAt(0);
    checkNotNull(testSheet, "Unable to get sheet 0 from the test Excel file");
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
  public void safeGetCellString_succeed() {
    ImmutableMap<ExcelCellIndex, String> expectedResults =
        ImmutableMap.<ExcelCellIndex, String>builder()
            .put(ExcelCellIndex.of(1, 1), "123")
            .put(ExcelCellIndex.of(1, 3), "abc")
            .put(ExcelCellIndex.of(3, 1), "4.567")
            // A cell in which text contains subscriptions and superscriptions.
            .put(ExcelCellIndex.of(3, 3), "abc,def,hij")
            // A cell in which text contains superscriptions at end.
            .put(ExcelCellIndex.of(4, 2), "abc")
            // A cell in which text using slightly different font color between the first
            // half and the second half. We expect to get a whole string here.
            .put(ExcelCellIndex.of(4, 3), "abcdef")
            .buildOrThrow();
    for (int row = 0; row < 5; row++) {
      for (int column = 0; column < 5; column++) {
        ExcelCellIndex idx = ExcelCellIndex.of(row, column);
        String expected = expectedResults.getOrDefault(idx, "");
        String actual = SheetParserUtils.safeGetCellString(testSheet, idx);
        assertEquals(
            String.format(
                "Unexpected cell value at %s, expected = '%s', actual = '%s'",
                idx.toString(), expected, actual),
            expected,
            actual);
      }
    }
  }

  @Test
  public void safeGetCellString_fail() {
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.safeGetCellString(null, ExcelCellIndex.of(0, 0)));
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.safeGetCellString(testSheet, ExcelCellIndex.of(-1, 0)));
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.safeGetCellString(testSheet, ExcelCellIndex.of(0, -1)));
    // Cell C6 (e.g., row = 5, column = 2) is a formula cell which is not supported.
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.safeGetCellString(testSheet, ExcelCellIndex.of(5, 2)));
  }

  @Test
  public void getColumnsWithTopBorder_succeed() {
    ImmutableSet<Integer> allColumns = ImmutableSet.of(0, 1, 2, 3, 4);

    // Row 0 has top border at all columns by definition.
    assertEquals(allColumns, SheetParserUtils.getColumnsWithTopBorder(testSheet, 0, allColumns));

    // Row 1 and 4 has top border at column 1, 2, 3, but not at other columns.
    assertEquals(
        ImmutableSet.of(1, 2, 3),
        SheetParserUtils.getColumnsWithTopBorder(testSheet, 1, allColumns));
    assertEquals(
        ImmutableSet.of(1, 2, 3),
        SheetParserUtils.getColumnsWithTopBorder(testSheet, 4, allColumns));

    // Row 2 has top border at column 3, but not at other columns.
    assertEquals(
        ImmutableSet.of(3), SheetParserUtils.getColumnsWithTopBorder(testSheet, 2, allColumns));

    // Row 3 has top border at column 1, but not at other columns.
    assertEquals(
        ImmutableSet.of(1), SheetParserUtils.getColumnsWithTopBorder(testSheet, 3, allColumns));

    // Row 5 has no top border at any column.
    assertEquals(
        ImmutableSet.of(), SheetParserUtils.getColumnsWithTopBorder(testSheet, 5, allColumns));
  }

  @Test
  public void getColumnsWithTopBorder_fail() {
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.getColumnsWithTopBorder(null, 0, ImmutableSet.of(0)));
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.getColumnsWithTopBorder(testSheet, -1, ImmutableSet.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.getColumnsWithTopBorder(testSheet, -1, ImmutableSet.of(0)));
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.getColumnsWithTopBorder(testSheet, 0, ImmutableSet.of(-1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            SheetParserUtils.getColumnsWithTopBorder(
                testSheet, 0, ImmutableSet.of(Integer.MAX_VALUE)));
  }

  @Test
  public void getRowsWithLeftBorder_expectedResult() {
    ImmutableSet<Integer> allRows = ImmutableSet.of(0, 1, 2, 3, 4);

    // Column 0 has left border at all rows by definition.
    assertEquals(allRows, SheetParserUtils.getRowsWithLeftBorder(testSheet, allRows, 0));

    // Column 1 and 4 has left border at row 1, 2, 3, but not at other rows.
    assertEquals(
        ImmutableSet.of(1, 2, 3), SheetParserUtils.getRowsWithLeftBorder(testSheet, allRows, 1));
    assertEquals(
        ImmutableSet.of(1, 2, 3), SheetParserUtils.getRowsWithLeftBorder(testSheet, allRows, 4));

    // Column 2 has left border at rows 3, but not at other rows.
    assertEquals(ImmutableSet.of(3), SheetParserUtils.getRowsWithLeftBorder(testSheet, allRows, 2));

    // Column 3 has left border at row 1, but not at other rows.
    assertEquals(ImmutableSet.of(1), SheetParserUtils.getRowsWithLeftBorder(testSheet, allRows, 3));

    // Column 5 has no left border at any row.
    assertEquals(ImmutableSet.of(), SheetParserUtils.getRowsWithLeftBorder(testSheet, allRows, 5));
  }

  @Test
  public void hasLeftBorder_exceptionalCases() {
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.getRowsWithLeftBorder(null, ImmutableSet.of(0), 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.getRowsWithLeftBorder(null, ImmutableSet.of(), 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.getRowsWithLeftBorder(testSheet, ImmutableSet.of(0), -1));
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetParserUtils.getRowsWithLeftBorder(testSheet, ImmutableSet.of(-1), 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            SheetParserUtils.getRowsWithLeftBorder(
                testSheet, ImmutableSet.of(Integer.MAX_VALUE), 0));
  }
}
