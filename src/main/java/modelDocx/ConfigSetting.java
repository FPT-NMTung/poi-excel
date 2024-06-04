package modelDocx;

import java.util.ArrayList;
import java.util.HashMap;

public class ConfigSetting {
    private ArrayList<CellConfig> generalData;
    private ArrayList<TableConfig> tableConfigs;

    public ConfigSetting() {
        this.generalData = new ArrayList<>();
        this.tableConfigs = new ArrayList<>();
    }

    public ConfigSetting(ArrayList<CellConfig> generalData, ArrayList<TableConfig> tableConfigs) {
        this.generalData = generalData;
        this.tableConfigs = tableConfigs;
    }

    public ArrayList<CellConfig> getGeneralData() {
        return generalData;
    }

    public void setGeneralData(ArrayList<CellConfig> generalData) {
        this.generalData = generalData;
    }

    public ArrayList<TableConfig> getTableConfigs() {
        return tableConfigs;
    }

    public void setTableConfigs(ArrayList<TableConfig> tableConfigs) {
        this.tableConfigs = tableConfigs;
    }
}
