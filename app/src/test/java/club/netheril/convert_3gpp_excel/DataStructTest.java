package club.netheril.convert_3gpp_excel;

import static org.junit.Assert.*;

import org.junit.Test;

public class DataStructTest {

  @Test
  public void createExcelCellIndex_succeed() {
    ExcelCellIndex a12 = ExcelCellIndex.of(11, 0);
    assertEquals("A12", a12.toString());

    ExcelCellIndex c99 = ExcelCellIndex.of(98, 2);
    assertEquals("C99", c99.toString());

    ExcelCellIndex ab1001 = ExcelCellIndex.of(1000, 27);
    assertEquals("AB1001", ab1001.toString());
  }

  @Test
  public void createExcelCellIndex_fail() {
    assertThrows(IllegalArgumentException.class, () -> ExcelCellIndex.of(-1, 0));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellIndex.of(0, -1));
  }

  @Test
  public void parseExcelCellName_succeed() {
    ExcelCellIndex a12 = ExcelCellIndex.of("A12");
    assertEquals(11, a12.row());
    assertEquals(0, a12.column());

    ExcelCellIndex c99 = ExcelCellIndex.of("C99");
    assertEquals(98, c99.row());
    assertEquals(2, c99.column());

    ExcelCellIndex ab1001 = ExcelCellIndex.of("AB1001");
    assertEquals(1000, ab1001.row());
    assertEquals(27, ab1001.column());
  }

  @Test
  public void parseExcelCellName_fail() {
    assertThrows(IllegalArgumentException.class, () -> ExcelCellIndex.of("12A"));
  }
}
