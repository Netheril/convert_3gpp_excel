package club.netheril.convert_3gpp_excel;

import org.junit.Test;
import static org.junit.Assert.*;

public class SheetParserUtilsTest {
    @Test
    public void parseExcelCellName_succeed() {
        CellIndex a12 = SheetParserUtils.parseExcelCellName("A12");
        assertEquals(11, a12.row.intValue());
        assertEquals(0, a12.col.intValue());

        CellIndex c99 = SheetParserUtils.parseExcelCellName("C99");
        assertEquals(98, c99.row.intValue());
        assertEquals(2, c99.col.intValue());

        CellIndex ab1001 = SheetParserUtils.parseExcelCellName("AB1001");
        assertEquals(1000, ab1001.row.intValue());
        assertEquals(27, ab1001.col.intValue());
    }

    @Test
    public void parseExcelCellName_fail() {
        assertThrows(
                IllegalArgumentException.class, () -> SheetParserUtils.parseExcelCellName("12A"));
    }
}
