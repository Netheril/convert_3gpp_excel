package club.netheril.convert_3gpp_excel;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.*;

import java.io.IOException;
import java.io.InputStream;

import com.google.common.collect.ImmutableList;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
        // We pick row [5, 20) from the Excel file in this test, which looks like this:
        // @formatter:off
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
        // @formatter:on
        TableMetadata metadata = new TableMetadata(
                "38.101-3", "h50",
                "5.3B.1.3-1", "",
                5, 20, 0, 7);

        TableRow dc7a = TableRow.fromColumns(ImmutableList.of(
                TableColumn.leaf(ImmutableList.of("DC_7A_n7A")),
                TableColumn.leaf(ImmutableList.of("DC_7A_n7A")),
                TableColumn.leaf(ImmutableList.of("5, 10, 15, 20")),
                TableColumn.leaf(ImmutableList.of("5, 10, 15, 20")),
                TableColumn.empty(),
                TableColumn.leaf(ImmutableList.of("40")),
                TableColumn.leaf(ImmutableList.of("0"))));

        TableRow dc41a = TableRow.fromColumns(ImmutableList.of(
                TableColumn.leaf(ImmutableList.of("DC_41A_n41A")),
                TableColumn.leaf(ImmutableList.of("DC_41A_n41A")),
                TableColumn.parent(ImmutableList.of(
                        TableRow.fromColumns(ImmutableList.of(
                                TableColumn.parent(ImmutableList.of(
                                        TableRow.fromColumns(ImmutableList.of(
                                                TableColumn.leaf(ImmutableList.of("20")),
                                                TableColumn.leaf(ImmutableList.of("40, 60, 80, 100")),
                                                TableColumn.empty())),
                                        TableRow.fromColumns(ImmutableList.of(
                                                TableColumn.empty(),
                                                TableColumn.leaf(ImmutableList.of("40, 60, 80, 100")),
                                                TableColumn.leaf(ImmutableList.of("20")))))),
                                TableColumn.leaf(ImmutableList.of("120")),
                                TableColumn.leaf(ImmutableList.of("0")))),
                        TableRow.fromColumns(ImmutableList.of(
                                TableColumn.parent(ImmutableList.of(
                                        TableRow.fromColumns(ImmutableList.of(
                                                TableColumn.leaf(ImmutableList.of("20")),
                                                TableColumn.leaf(ImmutableList.of("40, 50, 60, 80, 100")),
                                                TableColumn.empty())),
                                        TableRow.fromColumns(ImmutableList.of(
                                                TableColumn.empty(),
                                                TableColumn.leaf(ImmutableList.of("40, 50, 60, 80, 100")),
                                                TableColumn.leaf(ImmutableList.of("20")))))),
                                TableColumn.leaf(ImmutableList.of("120")),
                                TableColumn.leaf(ImmutableList.of("1")))),
                        TableRow.fromColumns(ImmutableList.of(
                                TableColumn.parent(ImmutableList.of(
                                        TableRow.fromColumns(ImmutableList.of(
                                                TableColumn.leaf(ImmutableList.of("20")),
                                                TableColumn.leaf(ImmutableList.of("10, 20, 30, 40, 50, 60")),
                                                TableColumn.empty())),
                                        TableRow.fromColumns(ImmutableList.of(
                                                TableColumn.empty(),
                                                TableColumn.leaf(ImmutableList.of("10, 20, 30, 40, 50, 60")),
                                                TableColumn.leaf(ImmutableList.of("20")))),
                                        TableRow.fromColumns(ImmutableList.of(
                                                TableColumn.leaf(ImmutableList.of("10")),
                                                TableColumn.leaf(ImmutableList.of("30, 40, 50, 60, 80, 100")),
                                                TableColumn.empty())),
                                        TableRow.fromColumns(ImmutableList.of(
                                                TableColumn.empty(),
                                                TableColumn.leaf(ImmutableList.of("30, 40, 50, 60, 80, 100")),
                                                TableColumn.leaf(ImmutableList.of("10")))))),
                                TableColumn.leaf(ImmutableList.of("120")),
                                TableColumn.leaf(ImmutableList.of("2"))))))));

        TableRow dc48d = TableRow.fromColumns(ImmutableList.of(
                TableColumn.leaf(ImmutableList.of("DC_48D_n48A")),
                TableColumn.leaf(ImmutableList.of("DC_48A_n48A")),
                TableColumn.parent(ImmutableList.of(
                        TableRow.fromColumns(ImmutableList.of(
                                TableColumn.leaf(ImmutableList.of("CA_48D_BCS0")),
                                TableColumn.leaf(ImmutableList.of("5, 10, 15, 20, 40")),
                                TableColumn.empty())),
                        TableRow.fromColumns(ImmutableList.of(
                                TableColumn.empty(),
                                TableColumn.leaf(ImmutableList.of("5, 10, 15, 20, 40")),
                                TableColumn.leaf(ImmutableList.of("CA_48D_BCS0")))))),
                TableColumn.leaf(ImmutableList.of("100")),
                TableColumn.leaf(ImmutableList.of("0"))));

        TableRow dc66c = TableRow.fromColumns(ImmutableList.of(
                TableColumn.leaf(ImmutableList.of("DC_66C_n66C")),
                TableColumn.parent(ImmutableList.of(
                        TableRow.fromColumns(ImmutableList.of(
                                TableColumn.leaf(ImmutableList.of("DC_66A_n66C")),
                                TableColumn.leaf(ImmutableList.of("5, 10, 15, 20")),
                                TableColumn.leaf(ImmutableList.of("5, 10, 15, 20, 40")),
                                TableColumn.empty(),
                                TableColumn.leaf(ImmutableList.of("50")),
                                TableColumn.leaf(ImmutableList.of("0")))),
                        TableRow.fromColumns(ImmutableList.of(
                                TableColumn.leaf(ImmutableList.of("DC_66C_n66A")),
                                TableColumn.parent(ImmutableList.of(
                                        TableRow.fromColumns(ImmutableList.of(
                                                TableColumn.leaf(ImmutableList.of("5, 10, 15, 20")),
                                                TableColumn.leaf(ImmutableList.of("5, 10, 15, 20, 25, 30, 40")),
                                                TableColumn.empty())),
                                        TableRow.fromColumns(ImmutableList.of(
                                                TableColumn.empty(),
                                                TableColumn.leaf(ImmutableList.of("5, 10, 15, 20, 25, 30, 40")),
                                                TableColumn.leaf(ImmutableList.of("5, 10, 15, 20")))))),
                                TableColumn.leaf(ImmutableList.of("60")),
                                TableColumn.leaf(ImmutableList.of("1"))))))));

        TableData expected = TableData.fromRows(ImmutableList.of(dc7a, dc41a, dc48d, dc66c));
        TableData actual = TableSheetParser.parse(testWorkbook, metadata);
        assertEquals(expected.rows().size(), actual.rows().size());
        for (int row = 0; row < expected.rows().size(); row++) {
            TableRow expectedRow = expected.rows().get(row);
            TableRow actualRow = actual.rows().get(row);
            assertEquals(String.format("Column size diff at row %d", row), expectedRow.columns().size(),
                    actualRow.columns().size());
            for (int column = 0; column < expectedRow.columns().size(); column++) {
                assertEquals(String.format("Diff at row %d column %d", row, column), expectedRow.columns().get(column),
                        actualRow.columns().get(column));
            }
        }
    }

    @Test
    public void parse_succeedOnWide() {
        // This is an additional test case which uses an "extremely wide" excel file, in
        // which one logical leaf column is composed of multiple physical Excel columns.
        // For example:
        // @formatter:off
        // ┌─────┬─────┬─────┬─────┬─────┐
        // │  F  │  G  │  H  │  I  │  J  │
        // ├─────┴─────┴─────┴─────┴─────┤
        // │            3 MHz            │
        // ├─────────────────────────────┤
        // │             Yes             │
        // └─────────────────────────────┘
        // @formatter:on
        TableMetadata metadata = new TableMetadata(
                "36.101", "h70",
                "5.6A.1-2", "",
                3, 11, 0, 32);

        TableRow ca1a3a = TableRow.fromColumns(
                ImmutableList.of(
                        TableColumn.leaf(ImmutableList.of("CA_1A-3A")),
                        TableColumn.leaf(ImmutableList.of("CA_1A-3A")),
                        TableColumn.parent(ImmutableList.of(
                                TableRow.fromColumns(ImmutableList.of(
                                        TableColumn.parent(ImmutableList.of(
                                                TableRow.fromColumns(ImmutableList.of(
                                                        TableColumn.leaf(ImmutableList.of("1")),
                                                        TableColumn.empty(),
                                                        TableColumn.empty(),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")))),
                                                TableRow.fromColumns(ImmutableList.of(
                                                        TableColumn.leaf(ImmutableList.of("3")),
                                                        TableColumn.empty(),
                                                        TableColumn.empty(),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")))))),
                                        TableColumn.leaf(ImmutableList.of("40")),
                                        TableColumn.leaf(ImmutableList.of("0")))),
                                TableRow.fromColumns(ImmutableList.of(
                                        TableColumn.parent(ImmutableList.of(
                                                TableRow.fromColumns(ImmutableList.of(
                                                        TableColumn.leaf(ImmutableList.of("1")),
                                                        TableColumn.empty(),
                                                        TableColumn.empty(),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")))),
                                                TableRow.fromColumns(ImmutableList.of(
                                                        TableColumn.leaf(ImmutableList.of("3")),
                                                        TableColumn.empty(),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                                        TableColumn.leaf(ImmutableList.of("Yes")))))),
                                        TableColumn.leaf(ImmutableList.of("40")),
                                        TableColumn.leaf(ImmutableList.of("1"))))))));

        TableRow ca1a1a3a = TableRow.fromColumns(
                ImmutableList.of(
                        TableColumn.leaf(ImmutableList.of("CA_1A-1A-3A")),
                        TableColumn.leaf(ImmutableList.of("-")),
                        TableColumn.parent(ImmutableList.of(
                                TableRow.fromColumns(ImmutableList.of(
                                        TableColumn.leaf(ImmutableList.of("1")),
                                        TableColumn.leaf(ImmutableList
                                                .of("See CA_1A-1A Bandwidth combination set 0 in Table 5.6A.1-3")))),
                                TableRow.fromColumns(ImmutableList.of(
                                        TableColumn.leaf(ImmutableList.of("3")),
                                        TableColumn.empty(),
                                        TableColumn.empty(),
                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                        TableColumn.leaf(ImmutableList.of("Yes")),
                                        TableColumn.leaf(ImmutableList.of("Yes"))))

                        )),
                        TableColumn.leaf(ImmutableList.of("60")),
                        TableColumn.leaf(ImmutableList.of("0"))));

        TableRow ca1a1a7c = TableRow.fromColumns(
                ImmutableList.of(
                        TableColumn.leaf(ImmutableList.of("CA_1A-1A-7C")),
                        TableColumn.leaf(ImmutableList.of("CA_7C")),
                        TableColumn.parent(ImmutableList.of(
                                TableRow.fromColumns(ImmutableList.of(
                                        TableColumn.leaf(ImmutableList.of("1")),
                                        TableColumn.leaf(ImmutableList
                                                .of("See CA_1A-1A Bandwidth Combination Set 0 in Table 5.6A.1-3")))),
                                TableRow.fromColumns(ImmutableList.of(
                                        TableColumn.leaf(ImmutableList.of("7")),
                                        TableColumn.leaf(ImmutableList
                                                .of("See CA_7C in Table 5.6A.1-1 of 36.101 Bandwidth combination set 2"))))

                        )),
                        TableColumn.leaf(ImmutableList.of("80")),
                        TableColumn.leaf(ImmutableList.of("0"))));

        TableData expected = TableData.fromRows(ImmutableList.of(ca1a3a, ca1a1a3a, ca1a1a7c));
        TableData actual = TableSheetParser.parse(testWorkbookWide, metadata);
        assertEquals(expected.rows().size(), actual.rows().size());
        for (int row = 0; row < expected.rows().size(); row++) {
            TableRow expectedRow = expected.rows().get(row);
            TableRow actualRow = actual.rows().get(row);
            assertEquals(String.format("Column size diff at row %d", row), expectedRow.columns().size(),
                    actualRow.columns().size());
            for (int column = 0; column < expectedRow.columns().size(); column++) {
                assertEquals(String.format("Diff at row %d column %d", row, column), expectedRow.columns().get(column),
                        actualRow.columns().get(column));
            }
        }
    }

    @Test
    public void parse_failureDueToInvalidBorder() {
        // We pick row [21, 26) from the Excel in this test, which looks like this:
        // @formatter:off
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
        // @formatter:on
        // It has an invalid cell border and thus can't be parsed.
        TableMetadata metadata = new TableMetadata(
                "38.101-3", "h50",
                "5.3B.1.3-1", "",
                21, 26, 0, 7);
        assertThrows(IllegalArgumentException.class, () -> TableSheetParser.parse(testWorkbook, metadata));
    }

    @Test
    public void parse_failureDueToInvalidDataRegion() {
        // We pick row [10, 20) from the Excel in this test, which is not a valid data
        // region.
        TableMetadata metadata = new TableMetadata(
                "38.101-3", "h50",
                "5.3B.1.3-1", "",
                10, 20, 0, 7);
        assertThrows(IllegalArgumentException.class, () -> TableSheetParser.parse(testWorkbook, metadata));
    }
}
