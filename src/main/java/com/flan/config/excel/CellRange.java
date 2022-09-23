package com.flan.config.excel;

public class CellRange {
    public final CellRef start;
    public final CellRef end;

    public CellRange(CellRef start, CellRef end) {
        if ( ! (start.col <= end.col && start.row <= end.row) )
            throw new RuntimeException("Bad cell range (start <= end): start: " + start + " end: " + end);

        this.start = start;
        this.end = end;
    }

    @Override
    public String toString() {
        return "CellRange{" +
                "start=" + start +
                ", end=" + end +
                '}';
    }

    public boolean inside(CellRef ref) {
        if ( start.row <= ref.row && end.row >= ref.row )
            if ( start.col <= ref.col && end.col >= ref.col)
                return true;
        return false;
    }
}
