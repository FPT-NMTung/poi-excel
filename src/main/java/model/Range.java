package model;

import org.apache.poi.ss.util.CellAddress;

public class Range {
    private String begin;
    private String end;
    private String[] columnData;

    public Range(String begin, String end, String[] columnData) {
        this.begin = begin;
        this.end = end;
        this.columnData = columnData;
    }

    public Range(String begin, String end) {
        this.begin = begin;
        this.end = end;
    }

    public String getBegin() {
        return begin;
    }

    public void setBegin(String begin) {
        this.begin = begin;
    }

    public String getEnd() {
        return end;
    }

    public void setEnd(String end) {
        this.end = end;
    }

    public String[] getColumnData() {
        return columnData;
    }

    public void setColumnData(String[] columnData) {
        this.columnData = columnData;
    }

    public boolean isColumnDataIsEmpty () {
        return columnData == null || columnData.length == 0 || (columnData.length == 1 && columnData[0].trim().isEmpty());
    }

    public int getHeightRange () {
        return new CellAddress(this.end).getRow() - new CellAddress(this.begin).getRow() + 1;
    }
}
