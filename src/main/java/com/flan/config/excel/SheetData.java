package com.flan.config.excel;

import java.util.ArrayList;
import java.util.HashMap;

public class SheetData {
    public final ArrayList<Row> rows = new ArrayList<>(10);
    public final HashMap<String, CellRange> tablesRanges = new HashMap<String, CellRange>();
}
