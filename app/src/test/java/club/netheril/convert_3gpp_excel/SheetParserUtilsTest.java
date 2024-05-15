package club.netheril.convert_3gpp_excel;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.*;

import com.google.common.collect.ImmutableMap;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import java.io.InputStream;
import java.io.IOException;

public class SheetParserUtilsTest {

    private final String TEST_EXCEL_FILE = "table_border_test_data.xlsx";

    private XSSFWorkbook testWorkbook;
    private XSSFSheet testSheet;

    @Before
    public void setUp() {
        InputStream file = getClass().getClassLoader().getResourceAsStream(TEST_EXCEL_FILE);
        if (file == null) {
            throw new RuntimeException(String.format("Unable to find file '%s'", TEST_EXCEL_FILE));
        }
        try {
            testWorkbook = new XSSFWorkbook(file);
        } catch (IOException | IllegalArgumentException e) {
            throw new RuntimeException(String.format("Unable to read Excel from file '%s'", TEST_EXCEL_FILE), e);
        }
        testSheet = testWorkbook.getSheetAt(0);
        if (testSheet == null) {
            throw new RuntimeException("Unable to get sheet 0 from the test Excel file.");
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
    public void parseExcelCellName_succeed() {
        ExcelCellIndex a12 = SheetParserUtils.parseExcelCellName("A12");
        assertEquals(11, a12.row());
        assertEquals(0, a12.col());

        ExcelCellIndex c99 = SheetParserUtils.parseExcelCellName("C99");
        assertEquals(98, c99.row());
        assertEquals(2, c99.col());

        ExcelCellIndex ab1001 = SheetParserUtils.parseExcelCellName("AB1001");
        assertEquals(1000, ab1001.row());
        assertEquals(27, ab1001.col());
    }

    @Test
    public void parseExcelCellName_fail() {
        assertThrows(
                IllegalArgumentException.class, () -> SheetParserUtils.parseExcelCellName("12A"));
    }

    @Test
    public void safeGetCellString_succeed() {
        ImmutableMap<ExcelCellIndex, String> expectedResults = ImmutableMap.<ExcelCellIndex, String>builder()
                .put(new ExcelCellIndex(1, 1), "123")
                .put(new ExcelCellIndex(1, 3), "abc")
                .put(new ExcelCellIndex(3, 1), "4.567")
                // A cell in which text contains subscriptions and superscriptions.
                .put(new ExcelCellIndex(3, 3), "abc,def,hij")
                // A cell in which text contains superscriptions at end.
                .put(new ExcelCellIndex(4, 2), "abc")
                // A cell in which text using slightly different font color between the first
                // half and the second half. We expect to get a whole string here.
                .put(new ExcelCellIndex(4, 3), "abcdef")
                .buildOrThrow();
        for (int row = 0; row < 5; row++) {
            for (int column = 0; column < 5; column++) {
                ExcelCellIndex idx = new ExcelCellIndex(row, column);
                String expected = expectedResults.getOrDefault(idx, "");
                String actual = SheetParserUtils.safeGetCellString(testSheet, idx);
                assertEquals(
                        String.format("Unexpected cell value at %s, expected = '%s', actual = '%s'",
                                idx.toString(), expected, actual),
                        expected, actual);
            }
        }
    }

    @Test
    public void safeGetCellString_fail() {
        assertThrows(
                IllegalArgumentException.class,
                () -> SheetParserUtils.safeGetCellString(null, new ExcelCellIndex(0, 0)));
        assertThrows(
                IllegalArgumentException.class,
                () -> SheetParserUtils.safeGetCellString(testSheet, new ExcelCellIndex(-1, 0)));
        assertThrows(
                IllegalArgumentException.class,
                () -> SheetParserUtils.safeGetCellString(testSheet, new ExcelCellIndex(0, -1)));
        // Cell C6 (e.g., row = 5, column = 2) is a formula cell which is not supported.
        assertThrows(
                IllegalArgumentException.class,
                () -> SheetParserUtils.safeGetCellString(testSheet, new ExcelCellIndex(5, 2)));
    }
}
