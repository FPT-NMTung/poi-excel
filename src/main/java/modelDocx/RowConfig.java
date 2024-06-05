package modelDocx;

import java.util.ArrayList;
import java.util.HashMap;

public class RowConfig {
    private String index;
    private int startRow;
    private int endRow;
    private RowConfig rowChildConfig;
    private ArrayList<CellConfig> mapCellConfig;

    public RowConfig(String index, int startRow, int endRow) {
        this.index = index;
        this.startRow = startRow;
        this.endRow = endRow;
        this.mapCellConfig =  new ArrayList<>();
    }

    public String getIndex() {
        return index;
    }

    public void setIndex(String index) {
        this.index = index;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }

    public RowConfig getRowChildConfig() {
        return rowChildConfig;
    }

    public void setRowChildConfig(RowConfig rowChildConfig) {
        this.rowChildConfig = rowChildConfig;
    }

    public ArrayList<CellConfig> getMapCellConfig() {
        return mapCellConfig;
    }

    public void setMapCellConfig(ArrayList<CellConfig> listCellConfig) {
        this.mapCellConfig = listCellConfig;
    }
}
