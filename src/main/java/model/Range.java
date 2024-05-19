package model;

public class Range {
    private String begin;
    private String end;
    private String[] columnData;

    public Range(String begin, String end, String[] columnData) {
        this.begin = begin;
        this.end = end;
        this.columnData = columnData;
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
}
