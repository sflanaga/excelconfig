package com.flan.config.excel;

public class TableData {
    public final String sheetname;
    public final String tablename;
    public final int height, width;
    public final int rowstart, colstart;
    public final String[][] data;

    public TableData(String sheetname, String tablename, CellRange range) {
        this.sheetname = sheetname;
        this.tablename = tablename;
        this.height = range.end.row - range.start.row +1;
        this.width = range.end.col - range.start.col +1;
        this.rowstart = range.start.row;
        this.colstart = range.start.col;
        this.data = new String[height][width];
    }

    public void mapSheet(Sheet sheet) {
        var sd = sheet.data().rows;
        for(var row: sd) {
            if ( row.row >= rowstart && row.row < rowstart + height) {
                int thisrow = row.row - rowstart;
                String[] datarow = null;
                try {
                    datarow = data[thisrow];
                } catch(ArrayIndexOutOfBoundsException e) {
                    throw new RuntimeException("mapping row for " + tablename + " bad mapping row: " + thisrow + " with height: " + height, e);
                }
                for(var cd: row.cols) {
                    if ( cd.col >= colstart && cd.col < colstart + width) {
                        int thiscol = cd.col - colstart;
                        try {
                            datarow[thiscol] = cd.val;
                        } catch(ArrayIndexOutOfBoundsException e) {
                            throw new RuntimeException("mapping column for " + tablename + " bad mapping: " +thisrow + ":" + thiscol + " with " + width + " by " + height , e);
                        }
                    }
                }
            }
        }
    }
}
