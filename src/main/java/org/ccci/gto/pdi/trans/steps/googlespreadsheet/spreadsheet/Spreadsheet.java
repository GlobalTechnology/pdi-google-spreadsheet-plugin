package org.ccci.gto.pdi.trans.steps.googlespreadsheet.spreadsheet;

import java.util.ArrayList;
import java.util.List;

public class Spreadsheet {

    public final List<SpreadsheetRow> rows;

    public Spreadsheet() {
        this.rows = new ArrayList<SpreadsheetRow>();
    }

    public int getRowCount() {
        return this.rows.size();
    }

    public int getColumnCount() {
        int columns = 0;
        for ( SpreadsheetRow row : this.rows ) {
            columns = Math.max( columns, row.getColumnCount() );
        }
        return columns;
    }

    public List<SpreadsheetCell> getCells() {
        List<SpreadsheetCell> cells = new ArrayList<SpreadsheetCell>();
        for ( int row = 1; row <= getRowCount(); row++ ) {
            for ( int col = 1; col <= getColumnCount(); col++ ) {
                SpreadsheetCell cell = rows.get( row - 1 ).cells.get( col - 1 );
                cell.row = row;
                cell.col = col;
                cell.id = String.format( "R%sC%s", row, col );
                cells.add( cell );
            }
        }
        return cells;
    }
}
