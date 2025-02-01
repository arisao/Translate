
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.sl.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import java.io.*;
import java.util.*;
import java.nio.file.*;

public class Translate {
    
    private static Map<String, String> replacementMap = new HashMap<>();

    public static void main(String[] args) {
        // 1. 置換ルールを読み込む
        loadReplacementRules("resources/replace_rules.csv");

        // 2. フォルダ内のすべてのWord, Excel, PowerPointファイルを処理
        String folderPath = "path";  // 変換対象のファイルがあるフォルダ
        processFilesInFolder(folderPath);
    }

    // 置換ルールのCSVを読み込む
    private static void loadReplacementRules(String filePath) {
        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(filePath), "UTF-8"))) {
            String line;
            while ((line = br.readLine()) != null) {
                String[] parts = line.split(",");
                if (parts.length == 2) {
                    replacementMap.put(parts[0].trim(), parts[1].trim());
                    System.out.println("読み込んだデータ: [" + parts[0] + "] -> [" + parts[1] + "]");
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    // フォルダ内のすべてのファイルを処理する
    private static void processFilesInFolder(String folderPath) {
        try {
            Files.walk(Paths.get(folderPath))
                 .filter(Files::isRegularFile)
                 .forEach(path -> processFile(path.toFile()));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // ファイルの種類を判定し、それぞれの処理を呼び出す
    private static void processFile(File file) {
        String fileName = file.getName();
        if (fileName.endsWith(".docx")) {
            processWordFile(file);
        } else if (fileName.endsWith(".xlsx")) {
            processExcelFile(file);
        } else if (fileName.endsWith(".pptx")) {
            processPowerPointFile(file);
        }
    }

    // Wordファイルを処理
    private static void processWordFile(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             XWPFDocument document = new XWPFDocument(fis)) {

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String text = paragraph.getText();
                
                // 置換処理
                for (Map.Entry<String, String> entry : replacementMap.entrySet()) {
                    text = text.replace(entry.getKey(), entry.getValue());                }

                // 古いランを削除
                while (paragraph.getRuns().size() > 0) {
                    paragraph.removeRun(0);
                }

                // 新しいランを追加
                XWPFRun newRun = paragraph.createRun();
                newRun.setText(text);

            }

            // 変更をファイルに保存
            try (FileOutputStream fos = new FileOutputStream(file)) {
                document.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Excelファイルを処理
    private static void processExcelFile(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
            
            for (Sheet sheet : workbook) {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                    	if (cell.getCellType() == CellType.STRING) {
                            String text = cell.getStringCellValue();
                            for (Map.Entry<String, String> entry : replacementMap.entrySet()) {
                                text = text.replace(entry.getKey(), entry.getValue());
                            }
                            cell.setCellValue(text);
                        }
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // PowerPointファイルを処理
    private static void processPowerPointFile(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             XMLSlideShow ppt = new XMLSlideShow(fis)) {
            
            for (XSLFSlide slide : ppt.getSlides()) {
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape textShape = (XSLFTextShape) shape;
                        String text = textShape.getText();
                        for (Map.Entry<String, String> entry : replacementMap.entrySet()) {
                            text = text.replace(entry.getKey(), entry.getValue());
                        }
                        textShape.setText(text);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(file)) {
                ppt.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
