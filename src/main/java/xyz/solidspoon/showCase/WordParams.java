package xyz.solidspoon.showCase;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class WordParams {
    private String tableOneName;
    private String tableTwoName;
    private String columnOne;
    private String columnTwo;
}
