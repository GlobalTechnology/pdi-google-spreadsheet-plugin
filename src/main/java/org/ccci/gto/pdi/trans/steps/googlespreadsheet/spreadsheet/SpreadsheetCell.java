package org.ccci.gto.pdi.trans.steps.googlespreadsheet.spreadsheet;

public class SpreadsheetCell {

    public final String value;
    public String id;
    public int row;
    public int col;

    public SpreadsheetCell( String value ) {
        this.value = value;
    }
}
