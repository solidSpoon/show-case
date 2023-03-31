package xyz.solidspoon.showCase;

import cn.hutool.core.annotation.AnnotationUtil;
import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.ReflectUtil;
import cn.hutool.core.util.StrUtil;
import org.apache.poi.xwpf.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class ContractGenerator {

    public static XWPFDocument generateWord2(Object paramEntity, String template, Map<Integer, List<Optional<Object>>> tableValues) {
        Field[] fields = ReflectUtil.getFields(paramEntity.getClass());
        Map<String, Object> param = Stream.of(fields)
                .collect(Collectors.toMap(Field::getName, field -> {
                    Object value = ReflectUtil.getFieldValue(paramEntity, field);
                    return MyUtil.nullBlank(value);
                }));
        return generateWord(param, template, tableValues);
    }


    /**
     * 根据指定的参数值、模板，生成 word 文档
     * 注意：其它模板需要根据情况进行调整
     *
     * @param param    变量集合
     * @param template 模板路径
     */
    public static XWPFDocument generateWord(Map<String, ?> param, String template, Map<Integer, List<Optional<Object>>> tableValues) {
        Map<String, String> processedParam = processParam(param);
        try (InputStream resourceStream = ContractGenerator.class.getClassLoader().getResourceAsStream(template)) {
            assert resourceStream != null;
            XWPFDocument doc = new XWPFDocument(resourceStream);
            List<XWPFParagraph> paragraphList = doc.getParagraphs();
            processParagraphs(paragraphList, processedParam);
            processTable(tableValues, doc, processedParam);
            return doc;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static Map<String, String> processParam(Map<String, ?> param) {
        Map<String, String> processedParam = new HashMap<>();
        param.forEach((key, value) -> processedParam.put(key, MyUtil.nullBlank(value)));
        return processedParam;
    }

    private static void processTable(Map<Integer, List<Optional<Object>>> tableValues, XWPFDocument doc, Map<String, String> param) {

        List<XWPFTable> tables = doc.getTables();
        for (XWPFTable table : tables) {
            List<XWPFTableRow> rows = table.getRows();
            for (XWPFTableRow row : rows) {
                List<XWPFTableCell> tableCells = row.getTableCells();
                for (XWPFTableCell cell : tableCells) {
                    List<XWPFParagraph> paragraphListTable = cell.getParagraphs();
                    processParagraphs(paragraphListTable, param);
                }
            }
        }
        tableValues.forEach((tableIndex, list) -> {
            XWPFTable xwpfTable = tables.get(tableIndex);
            List<XWPFTableRow> rows = xwpfTable.getRows();
            int targetSize = list.size();
            int sourceSize = rows.size();
            if (targetSize > sourceSize) {
                XWPFTableRow sourceRow = rows.get(sourceSize - 1);
                for (int j = 0; j < targetSize - sourceSize; j++) {
                    copy(xwpfTable, sourceRow, sourceSize + j);
                }
            }
        });
        for (Map.Entry<Integer, List<Optional<Object>>> integerListEntry : tableValues.entrySet()) {
            int tableIndex = integerListEntry.getKey();
            if (tableIndex >= tables.size()) {
                throw new RuntimeException("表格数量不够");
            }
            XWPFTable table = tables.get(tableIndex);
            List<Optional<Object>> tableList = integerListEntry.getValue();
            if (CollUtil.isEmpty(tableList)) {
                continue;
            }
            fillTable(tableList, table);
        }

    }

    /**
     * 拷贝赋值行
     */
    public static void copy(XWPFTable table, XWPFTableRow sourceRow, int rowIndex) {
        // 在表格指定位置新增一行
        XWPFTableRow targetRow = table.insertNewTableRow(rowIndex);
        // 复制行属性
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        List<XWPFTableCell> cellList = sourceRow.getTableCells();
        if (null == cellList) {
            return;
        }
        // 复制列及其属性和内容
        XWPFTableCell targetCell;
        for (XWPFTableCell sourceCell : cellList) {
            targetCell = targetRow.addNewTableCell();
            // 列属性
            targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            // 段落属性
            if (sourceCell.getParagraphs() != null && sourceCell.getParagraphs().size() > 0) {
                targetCell.getParagraphs().get(0).getCTP().setPPr(sourceCell.getParagraphs().get(0).getCTP().getPPr());
                if (sourceCell.getParagraphs().get(0).getRuns() != null && sourceCell.getParagraphs().get(0).getRuns().size() > 0) {
                    XWPFRun cellR = targetCell.getParagraphs().get(0).createRun();
                    cellR.setText(sourceCell.getText());
                    cellR.setBold(sourceCell.getParagraphs().get(0).getRuns().get(0).isBold());
                } else {
                    targetCell.setText(sourceCell.getText());
                }
            } else {
                targetCell.setText(sourceCell.getText());
            }
        }
    }

    private static <Object> void fillTable(List<Optional<Object>> tableList, XWPFTable table) {
        List<XWPFTableRow> rows = table.getRows();
        int row = 0;
        for (Optional<Object> t : tableList) {
            if (row >= rows.size()) {
                throw new RuntimeException("表格行数不够");
            }
            XWPFTableRow xwpfTableRow = rows.get(row);
            t.ifPresent(to -> fillTableRow(to, xwpfTableRow));
            row++;
        }
    }


    private static <T> void fillTableRow(T table, XWPFTableRow row) {
        for (Field field : ReflectUtil.getFields(table.getClass())) {
            if (!AnnotationUtil.hasAnnotation(field, TableColumnIndex.class)) {
                continue;
            }
            TableColumnIndex annotation = field.getAnnotation(TableColumnIndex.class);
            int index = Integer.parseInt(annotation.value());
            String value = Optional.ofNullable(ReflectUtil.getFieldValue(table, field))
                    .orElse("").toString();
            row.getCell(index).setText(value);
        }
    }

    /**
     * 处理段落
     */
    @SuppressWarnings({"unused", "rawtypes"})
    public static void processParagraphs(List<XWPFParagraph> paragraphList, Map<String, String> params) {
        if (CollUtil.isEmpty(paragraphList)) {
            return;
        }
        List<List<XWPFRun>> collect = paragraphList.stream()
                .map(ContractGenerator::mergeRuns)
                .map(XWPFParagraph::getRuns)
                .toList();
        List<XWPFRun> xwpfRuns = collect.stream()
                .flatMap(List::stream)
                .filter(run -> StrUtil.isNotBlank(run.getText(0)))
                .toList();
        xwpfRuns
                .forEach(run -> processIfRunInParam(params, run));
    }

    /**
     * 合并相同样式的run, 避免匹配不到占位符
     *
     * @param paragraph
     * @return
     */
    private static XWPFParagraph mergeRuns(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs.size() < 2) {
            return paragraph;
        }
        XWPFRun preRun = runs.get(0);
        for (int i = 1; i < runs.size(); i++) {
            XWPFRun run = runs.get(i);
            if (StrUtil.isBlank(run.getText(0))) {
                continue;
            }
            if (isSameStyle(preRun, run)) {
                preRun.setText(preRun.getText(0) + run.getText(0), 0);
                paragraph.removeRun(i);
                i--;
            } else {
                preRun = run;
            }
        }
        return paragraph;
    }

    private static boolean isSameStyle(XWPFRun preRun, XWPFRun run) {
        boolean equals;
        try {
            equals = run.getCTR().getRPr().getRFonts().getAscii().equals(preRun.getCTR().getRPr().getRFonts().getAscii());
        } catch (NullPointerException ignored) {
            equals = true;
        }
        return run.isBold() == preRun.isBold()
                && run.isItalic() == preRun.isItalic()
                && run.isStrikeThrough() == preRun.isStrikeThrough()
                && run.getFontSize() == preRun.getFontSize()
                && Objects.equals(run.getColor(), preRun.getColor())
                && run.getUnderline() == preRun.getUnderline()
//                && run.getVerticalAlignment().equals(preRun.getVerticalAlignment())
                && equals;
    }

    private static void processIfRunInParam(Map<String, String> param, XWPFRun run) {
        String text = run.getText(0);
        Optional<Map.Entry<String, String>> paramEntity = param.entrySet().stream()
                .filter(entry -> text.contains(toKey(entry.getKey()))).findAny();
        if (paramEntity.isEmpty()) {
            return;
        }
        Map.Entry<String, String> entry = paramEntity.get();
        String key = entry.getKey();
        String value = Optional.ofNullable(entry.getValue()).orElse("");
        run.setText(text.replace(toKey(key), value), 0);
    }

    private static String toKey(String key) {
        return "${" + key + "}";
    }
}
