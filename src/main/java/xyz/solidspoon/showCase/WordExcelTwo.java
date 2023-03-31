package xyz.solidspoon.showCase;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class WordExcelTwo {
    @TableColumnIndex("0")
    private String columnOne;
    @TableColumnIndex("1")
    private String columnTwo;
}
