package model;

public class ConfigSetting {
    private int totalGroup;
    private boolean isHasGeneralData;
    private boolean isMergeCell;
    private Range[] arrRange;

    public ConfigSetting(int totalGroup, boolean isHasGeneralData, boolean isMergeCell) {
        this.totalGroup = totalGroup;
        this.arrRange = new Range[totalGroup];
        this.isHasGeneralData = isHasGeneralData;
        this.isMergeCell = isMergeCell;
    }

    public int getTotalGroup() {
        return totalGroup;
    }

    public void setTotalGroup(int totalGroup) {
        this.totalGroup = totalGroup;
    }

    public Range[] getArrRange() {
        return arrRange;
    }

    public void setArrRange(Range[] arrRange) {
        this.arrRange = arrRange;
    }

    public boolean isHasGeneralData() {
        return isHasGeneralData;
    }

    public void setHasGeneralData(boolean hasGeneralData) {
        isHasGeneralData = hasGeneralData;
    }

    public boolean isMergeCell() {
        return isMergeCell;
    }

    public void setMergeCell(boolean mergeCell) {
        isMergeCell = mergeCell;
    }
}