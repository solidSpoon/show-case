package xyz.solidspoon.showCase;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class WordExcelOne {
    @TableColumnIndex("0")
    private String columnOne;
    @TableColumnIndex("1")
    private String columnTwo;
    @TableColumnIndex("2")
    private String columnThree;
}
