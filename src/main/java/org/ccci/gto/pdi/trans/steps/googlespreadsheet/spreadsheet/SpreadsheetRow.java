package org.ccci.gto.pdi.trans.steps.googlespreadsheet.spreadsheet;

import java.util.ArrayList;
import java.util.List;

public class SpreadsheetRow {
    public final List<SpreadsheetCell> cells;

    public SpreadsheetRow() {
        this.cells = new ArrayList<SpreadsheetCell>();
    }

    public int getColumnCount() {
        return this.cells.size();
    }
}
