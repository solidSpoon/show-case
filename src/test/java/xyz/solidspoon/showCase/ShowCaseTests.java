package xyz.solidspoon.showCase;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Optional;

@SpringBootTest
class ShowCaseTests {

    @Test
    void fillExcel() {
        WordParams wordParams = WordParams.builder()
                .tableOneName("TableOne")
                .tableTwoName("TableTwo")
                .columnOne("ColumnOne")
                .columnTwo("ColumnTwo")
                .build();
        List<Optional<Object>> tableOneRows = generateTableOneRows();
        List<Optional<Object>> tableTwoRows = generateTableTwoRows();
        Map<Integer, List<Optional<Object>>> tableRows = Map.of(
                0, tableOneRows,
                1, tableTwoRows
        );
        XWPFDocument xwpfDocument = TemplateFiller.generateWord2(wordParams, "template.docx", tableRows);
        //保存到本地
        saveDocumentToFile(xwpfDocument, "output.docx");

    }

    private void saveDocumentToFile(XWPFDocument xwpfDocument, String fileName) {
        // 获取项目路径
        String projectPath = System.getProperty("user.dir");

        // 设置输出文件路径
        File outputFile = new File(projectPath, fileName);

        // 将文档写入文件
        try (FileOutputStream out = new FileOutputStream(outputFile)) {
            xwpfDocument.write(out);
        } catch (IOException e) {
            System.err.println("保存文档时出现错误: " + e.getMessage());
        }

    }

    private List<Optional<Object>> generateTableTwoRows() {
        WordExcelTwo build = WordExcelTwo.builder()
                .columnOne("One")
                .columnTwo("Two")
                .build();
        return List.of(Optional.empty(), Optional.of(build),Optional.of(build));
    }

    private static List<Optional<Object>> generateTableOneRows() {
        WordExcelOne build = WordExcelOne.builder()
                .columnOne("One")
                .columnTwo("Two")
                .columnThree("Three")
                .build();
        return List.of(Optional.empty(), Optional.of(build),Optional.of(build));
    }


}
