package model;

public class ConfigSetting {
    private int totalGroup;
    private Range[] arrRange;

    public ConfigSetting(int totalGroup) {
        this.totalGroup = totalGroup;
        this.arrRange = new Range[totalGroup];
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
}