package com.flan.config.excel;

final public class CellRef {

    public final int row;
    public final int col;

    public CellRef(int row, int col) {
        this.row = row;
        this.col = col;
    }

    public CellRef(String s) {
        int i = 0;
        for (i = 0; i < s.length(); i++) {
            if ( Character.isDigit(s.charAt(i)) )
                break;
        }
        String digs = s.substring(i);
        this.col = lettersToRow(s.substring(0, i));
        this.row = Integer.valueOf(digs);
    }

    public static int lettersToRow(String s) {
        int v = 0;
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            v *= 26;
            if ( c >= 'A' && c <='Z' )
                v += ((int)s.charAt(i)) - ((int)'A') + 1;
            else
                throw new RuntimeException("Bad cell reference letters: " + s);
        }
        return v;
    }

    @Override
    public String toString() {
        return "CellRef{" +
                "row=" + row +
                ", col=" + col +
                '}';
    }

    public static void main(String[] args) {
        for(String s: new String[] { /* "A1", "B1", */ "DC2020"}) {
            System.out.println("s: " + s + " = " + new CellRef(s));
        }
    }
}
