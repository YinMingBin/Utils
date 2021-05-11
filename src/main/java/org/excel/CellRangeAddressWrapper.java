package org.excel;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Objects;

/**
 * @author Administrator
 */
public class CellRangeAddressWrapper implements Comparable<CellRangeAddressWrapper> {

    public CellRangeAddress range;

    public CellRangeAddressWrapper(CellRangeAddress theRange) {
          this.range = theRange;
    }

    @Override
    public int compareTo(CellRangeAddressWrapper craw) {
        if (range.getFirstColumn() < craw.range.getFirstColumn()
                    && range.getFirstRow() < craw.range.getFirstRow()) {
              return -1;
        } else if (range.getFirstColumn() == craw.range.getFirstColumn()
                    && range.getFirstRow() == craw.range.getFirstRow()) {
              return 0;
        } else {
              return 1;
        }
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {return true;}
        if (o == null || getClass() != o.getClass()) {return false;}
        CellRangeAddressWrapper that = (CellRangeAddressWrapper) o;
        return Objects.equals(range, that.range);
    }

    @Override
    public int hashCode() {
        return Objects.hash(range);
    }
}
