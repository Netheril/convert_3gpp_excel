package club.netheril.convert_3gpp_excel;

final class ExcelCellIndex {
    private int rowValue;
    private int colValue;

    public ExcelCellIndex(int row, int col) {
        this.rowValue = row;
        this.colValue = col;
    }

    public int row() {
        return rowValue;
    }

    public int col() {
        return colValue;
    }

    @Override
    public String toString() {
        return String.format("(%d, %d)", rowValue, colValue);
    }

    @Override
    public int hashCode() {
        return rowValue * 65536 + colValue;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null || getClass() != obj.getClass()) {
            return false;
        }
        ExcelCellIndex other = (ExcelCellIndex) obj;
        return row() == other.row() && col() == other.col();
    }
}
