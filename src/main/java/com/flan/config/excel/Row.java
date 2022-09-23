package com.flan.config.excel;

import java.util.ArrayList;

final public class Row {
    public final int row;
    public final ArrayList<CellData> cols;

    public Row(int row) {
        this.row = row;
        cols = new ArrayList<>(10);
    }
}
