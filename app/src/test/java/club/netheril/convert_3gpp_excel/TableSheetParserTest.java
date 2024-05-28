package club.netheril.convert_3gpp_excel;

import static org.junit.Assert.*;

import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class TableSheetParserTest {
  private XSSFWorkbook testWorkbook;
  private XSSFWorkbook testWorkbookWide;

  private XSSFWorkbook openExcelFile(String fileName) {
    InputStream file = getClass().getClassLoader().getResourceAsStream(fileName);
    if (file == null) {
      throw new RuntimeException(String.format("Unable to find file '%s'", fileName));
    }
    try {
      return new XSSFWorkbook(file);
    } catch (IOException | IllegalArgumentException e) {
      throw new RuntimeException(String.format("Unable to read Excel from file '%s'", fileName), e);
    }
  }

  @Before
  public void setUp() {
    testWorkbook = openExcelFile("table_5.3B.1.3-1.xlsx");
    testWorkbookWide = openExcelFile("table_5.6A.1-2.xlsx");
  }

  @After
  public void tearDown() {
    try {
      testWorkbook.close();
      testWorkbookWide.close();
    } catch (IOException e) {
      throw new RuntimeException("Unable to close test workbooks");
    }
  }

  @Test
  public void parse_succeed() {

    TableMetadata metadata =
        TableMetadata.builder()
            .setSpecName("38.101-3")
            .setSpecVersion("h50")
            .setTableSerialNumber("5.3B.1.3-1")
            .setTableTitle("")
            .setTableDataRect(ExcelRect.of("A6", "G20"))
            .build();

    // We pick row [5, 20) from the Excel file in this test, which looks like this:
    // ┌─────────────┬─────────────┬─────────────┬─────────────────────┬─────────────┬─────┬───┐
    // │ DC_7A_n7A   │ DC_7A_n7A   │ 5,10,15,20  │ 5,10,15,20          │             │ 40  │ 0 │
    // ├─────────────┼─────────────┼─────────────┼─────────────────────┼─────────────┼─────┼───┤
    // │             │             │ 20          │ 40,60,80,100        │             │     │   │
    // │             │             ├─────────────┼─────────────────────┼─────────────┤ 120 │ 0 │
    // │             │             │             │ 40,60,80,100        │ 20          │     │   │
    // │             │             ├─────────────┼─────────────────────┼─────────────┼─────┼───┤
    // │             │             │ 20          │ 40,50,60,80,100     │             │     │   │
    // │             │             ├─────────────┼─────────────────────┼─────────────┤ 120 │ 1 │
    // │             │             │             │ 40,50,60,80,100     │ 20          │     │   │
    // │ DC_41A_n41A │ DC_41A_n41A ├─────────────┼─────────────────────┼─────────────┼─────┼───┤
    // │             │             │ 20          │ 10,20,30,40,50,60   │             │     │   │
    // │             │             ├─────────────┼─────────────────────┼─────────────┤     │   │
    // │             │             │             │ 10,20,30,40,50,60   │ 20          │     │   │
    // │             │             ├─────────────┼─────────────────────┼─────────────┤ 120 │ 2 │
    // │             │             │ 10          │ 30,40,50,60,80,100  │             │     │   │
    // │             │             ├─────────────┼─────────────────────┼─────────────┤     │   │
    // │             │             │             │ 30,40,50,60,80,100  │ 10          │     │   │
    // ├─────────────┼─────────────┼─────────────┼─────────────────────┼─────────────┼─────┼───┤
    // │             │             │ CA_48D_BCS0 │ 5,10,15,20,40       │             │     │   │
    // │ DC_48D_n48A │ DC_48A_n48A ├─────────────┼─────────────────────┼─────────────┤ 100 │ 0 │
    // │             │             │             │ 5,10,15,20,40       │ CA_48D_BCS0 │     │   │
    // ├─────────────┼─────────────┼─────────────┼─────────────────────┼─────────────┼─────┼───┤
    // │             │ DC_66A_n66C │ 5,10,15,20  │ 5,10,15,20,40       │             │ 50  │ 0 │
    // │             ├─────────────┼─────────────┼─────────────────────┼─────────────┼─────┼───┤
    // │ DC_66C_n66C │             │ 5,10,15,20  │ 5,10,15,20,25,30,40 │             │     │   │
    // │             │ DC_66C_n66A ├─────────────┼─────────────────────┼─────────────┤ 60  │ 1 │
    // │             │             │             │ 5,10,15,20,25,30,40 │ 5,10,15,20  │     │   │
    // └─────────────┴─────────────┴─────────────┴─────────────────────┴─────────────┴─────┴───┘

    TableRow dc7a =
        TableRow.of(
            TableColumn.leaf("DC_7A_n7A"),
            TableColumn.leaf("DC_7A_n7A"),
            TableColumn.leaf("5, 10, 15, 20"),
            TableColumn.leaf("5, 10, 15, 20"),
            TableColumn.empty(),
            TableColumn.leaf("40"),
            TableColumn.leaf("0"));

    TableRow dc41a =
        TableRow.of(
            TableColumn.leaf("DC_41A_n41A"),
            TableColumn.leaf("DC_41A_n41A"),
            TableColumn.parent(
                TableRow.of(
                    TableColumn.parent(
                        TableRow.of(
                            TableColumn.leaf("20"),
                            TableColumn.leaf("40, 60, 80, 100"),
                            TableColumn.empty()),
                        TableRow.of(
                            TableColumn.empty(),
                            TableColumn.leaf("40, 60, 80, 100"),
                            TableColumn.leaf("20"))),
                    TableColumn.leaf("120"),
                    TableColumn.leaf("0")),
                TableRow.of(
                    TableColumn.parent(
                        TableRow.of(
                            TableColumn.leaf("20"),
                            TableColumn.leaf("40, 50, 60, 80, 100"),
                            TableColumn.empty()),
                        TableRow.of(
                            TableColumn.empty(),
                            TableColumn.leaf("40, 50, 60, 80, 100"),
                            TableColumn.leaf("20"))),
                    TableColumn.leaf("120"),
                    TableColumn.leaf("1")),
                TableRow.of(
                    TableColumn.parent(
                        TableRow.of(
                            TableColumn.leaf("20"),
                            TableColumn.leaf("10, 20, 30, 40, 50, 60"),
                            TableColumn.empty()),
                        TableRow.of(
                            TableColumn.empty(),
                            TableColumn.leaf("10, 20, 30, 40, 50, 60"),
                            TableColumn.leaf("20")),
                        TableRow.of(
                            TableColumn.leaf("10"),
                            TableColumn.leaf("30, 40, 50, 60, 80, 100"),
                            TableColumn.empty()),
                        TableRow.of(
                            TableColumn.empty(),
                            TableColumn.leaf("30, 40, 50, 60, 80, 100"),
                            TableColumn.leaf("10"))),
                    TableColumn.leaf("120"),
                    TableColumn.leaf("2"))));

    TableRow dc48d =
        TableRow.of(
            TableColumn.leaf("DC_48D_n48A"),
            TableColumn.leaf("DC_48A_n48A"),
            TableColumn.parent(
                TableRow.of(
                    TableColumn.leaf("CA_48D_BCS0"),
                    TableColumn.leaf("5, 10, 15, 20, 40"),
                    TableColumn.empty()),
                TableRow.of(
                    TableColumn.empty(),
                    TableColumn.leaf("5, 10, 15, 20, 40"),
                    TableColumn.leaf("CA_48D_BCS0"))),
            TableColumn.leaf("100"),
            TableColumn.leaf("0"));

    TableRow dc66c =
        TableRow.of(
            TableColumn.leaf("DC_66C_n66C"),
            TableColumn.parent(
                TableRow.of(
                    TableColumn.leaf("DC_66A_n66C"),
                    TableColumn.leaf("5, 10, 15, 20"),
                    TableColumn.leaf("5, 10, 15, 20, 40"),
                    TableColumn.empty(),
                    TableColumn.leaf("50"),
                    TableColumn.leaf("0")),
                TableRow.of(
                    TableColumn.leaf("DC_66C_n66A"),
                    TableColumn.parent(
                        TableRow.of(
                            TableColumn.leaf("5, 10, 15, 20"),
                            TableColumn.leaf("5, 10, 15, 20, 25, 30, 40"),
                            TableColumn.empty()),
                        TableRow.of(
                            TableColumn.empty(),
                            TableColumn.leaf("5, 10, 15, 20, 25, 30, 40"),
                            TableColumn.leaf("5, 10, 15, 20"))),
                    TableColumn.leaf("60"),
                    TableColumn.leaf("1"))));

    TableData expected = TableData.of(dc7a, dc41a, dc48d, dc66c);
    TableData actual = TableSheetParser.parse(testWorkbook, metadata);
    assertEquals(expected.rows().size(), actual.rows().size());
    for (int row = 0; row < expected.rows().size(); row++) {
      TableRow expectedRow = expected.rows().get(row);
      TableRow actualRow = actual.rows().get(row);
      assertEquals(
          String.format("Column size diff at row %d", row),
          expectedRow.columns().size(),
          actualRow.columns().size());
      for (int column = 0; column < expectedRow.columns().size(); column++) {
        assertEquals(
            String.format("Diff at row %d column %d", row, column),
            expectedRow.columns().get(column),
            actualRow.columns().get(column));
      }
    }
  }

  @Test
  public void parse_succeedOnWide() {
    TableMetadata metadata =
        TableMetadata.builder()
            .setSpecName("36.101")
            .setSpecVersion("h70")
            .setTableSerialNumber("5.6A.1-2")
            .setTableTitle("")
            .setTableDataRect(ExcelRect.of("A4", "AF11"))
            .build();

    // This is an additional test case which uses an "extremely wide" excel file, in
    // which one logical leaf column is composed of multiple physical Excel columns.
    // For example:
    // ┌─────┬─────┬─────┬─────┬─────┐
    // │  F  │  G  │  H  │  I  │  J  │
    // ├─────┴─────┴─────┴─────┴─────┤
    // │            3 MHz            │
    // ├─────────────────────────────┤
    // │             Yes             │
    // └─────────────────────────────┘

    TableRow ca1a3a =
        TableRow.of(
            TableColumn.leaf("CA_1A-3A"),
            TableColumn.leaf("CA_1A-3A"),
            TableColumn.parent(
                TableRow.of(
                    TableColumn.parent(
                        TableRow.of(
                            TableColumn.leaf("1"),
                            TableColumn.empty(),
                            TableColumn.empty(),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes")),
                        TableRow.of(
                            TableColumn.leaf("3"),
                            TableColumn.empty(),
                            TableColumn.empty(),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"))),
                    TableColumn.leaf("40"),
                    TableColumn.leaf("0")),
                TableRow.of(
                    TableColumn.parent(
                        TableRow.of(
                            TableColumn.leaf("1"),
                            TableColumn.empty(),
                            TableColumn.empty(),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes")),
                        TableRow.of(
                            TableColumn.leaf("3"),
                            TableColumn.empty(),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"),
                            TableColumn.leaf("Yes"))),
                    TableColumn.leaf("40"),
                    TableColumn.leaf("1"))));

    TableRow ca1a1a3a =
        TableRow.of(
            TableColumn.leaf("CA_1A-1A-3A"),
            TableColumn.leaf("-"),
            TableColumn.parent(
                TableRow.of(
                    TableColumn.leaf("1"),
                    TableColumn.leaf(
                        "See CA_1A-1A Bandwidth combination set 0 in Table" + " 5.6A.1-3")),
                TableRow.of(
                    TableColumn.leaf("3"),
                    TableColumn.empty(),
                    TableColumn.empty(),
                    TableColumn.leaf("Yes"),
                    TableColumn.leaf("Yes"),
                    TableColumn.leaf("Yes"),
                    TableColumn.leaf("Yes"))),
            TableColumn.leaf("60"),
            TableColumn.leaf("0"));

    TableRow ca1a1a7c =
        TableRow.of(
            TableColumn.leaf("CA_1A-1A-7C"),
            TableColumn.leaf("CA_7C"),
            TableColumn.parent(
                TableRow.of(
                    TableColumn.leaf("1"),
                    TableColumn.leaf(
                        "See CA_1A-1A Bandwidth Combination Set 0 in Table" + " 5.6A.1-3")),
                TableRow.of(
                    TableColumn.leaf("7"),
                    TableColumn.leaf(
                        "See CA_7C in Table 5.6A.1-1 of 36.101 Bandwidth" + " combination set 2"))),
            TableColumn.leaf("80"),
            TableColumn.leaf("0"));

    TableData expected = TableData.of(ca1a3a, ca1a1a3a, ca1a1a7c);
    TableData actual = TableSheetParser.parse(testWorkbookWide, metadata);
    assertEquals(expected.rows().size(), actual.rows().size());
    for (int row = 0; row < expected.rows().size(); row++) {
      TableRow expectedRow = expected.rows().get(row);
      TableRow actualRow = actual.rows().get(row);
      assertEquals(
          String.format("Column size diff at row %d", row),
          expectedRow.columns().size(),
          actualRow.columns().size());
      for (int column = 0; column < expectedRow.columns().size(); column++) {
        assertEquals(
            String.format("Diff at row %d column %d", row, column),
            expectedRow.columns().get(column),
            actualRow.columns().get(column));
      }
    }
  }

  @Test
  public void parse_failureDueToInvalidBorder() {
    // We pick row [21, 26) from the Excel in this test, which looks like this:
    // ┌─────────────┬─────────────┬─────────────┬─────────────────────┬─────────────┬─────┬───┐
    // │ DC_7A_n7A   │ DC_7A_n7A   │ 5,10,15,20  │ 5,10,15,20          │             │ 40  │ 0 │
    // ├─────────────┼─────────────┼─────────────┼─────────────────────┼─────────────┼─────┼───┤
    // │             │             │ 20          │ 40,60,80,100        │             │     │   │
    // │             │             ├─────────────┼─────────────────────┼─────────────┤ 120 │ 0 │
    // │             │             │             │ 40,60,80,100        │ 20          │     │   │
    // │ DC_41A_n41A │ DC_41A_n41A ├─────────────┼─────────────────────┼─────────────┼─────┼───┤
    // │             │             │ 20          │ 40,50,60,80,100     │             │     │   │
    // │             │             ├─────────────┼─────────────────────┤             │     │   │
    // │             │             │             │ 40,50,60,80,100     │ 20          │ 120 │ 1 │
    // │             │             │             ├─────────────────────┼─────────────┤     │   │
    // │             │             │             │ 40,50,60,80,100     │ 20          │     │   │
    // └─────────────┴─────────────┴─────────────┴─────────────────────┴─────────────┴─────┴───┘
    // It has an invalid cell border and thus can't be parsed.

    TableMetadata metadata =
        TableMetadata.builder()
            .setSpecName("38.101-3")
            .setSpecVersion("h50")
            .setTableSerialNumber("5.3B.1.3-1")
            .setTableTitle("")
            .setTableDataRect(ExcelRect.of(ExcelCellIndex.of(21, 0), ExcelCellIndex.of(26, 7)))
            .build();
    assertThrows(
        IllegalArgumentException.class, () -> TableSheetParser.parse(testWorkbook, metadata));
  }

  @Test
  public void parse_failureDueToInvalidDataRegion() {
    // We pick row [10, 20) from the Excel in this test, which is not a valid data
    // region.
    TableMetadata metadata =
        TableMetadata.builder()
            .setSpecName("38.101-3")
            .setSpecVersion("h50")
            .setTableSerialNumber("5.3B.1.3-1")
            .setTableTitle("")
            .setTableDataRect(ExcelRect.of(ExcelCellIndex.of(10, 0), ExcelCellIndex.of(20, 7)))
            .build();
    assertThrows(
        IllegalArgumentException.class, () -> TableSheetParser.parse(testWorkbook, metadata));
  }
}
