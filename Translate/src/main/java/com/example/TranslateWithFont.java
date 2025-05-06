package com.example;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.logging.log4j.core.tools.picocli.CommandLine.Help.TextTable.Cell;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.LocaleID;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xslf.usermodel.XSLFTheme;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.impl.xb.xsdschema.ListDocument.List;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontScheme;
import org.openxmlformats.schemas.drawingml.x2006.main.CTOfficeStyleSheet;

public class TranslateWithFont {
	//マップを宣言
    private static Map<String, String> replacementMap = new HashMap<>();
    //任意のフォント名を宣言
    private static final String FONT_NAME = "BIZ UDGothic";

    public static void main(String[] args) {
        if (args.length < 2) {
            System.out.println("使い方: java TranslateWithFont <置換CSVファイルパス> <変換対象フォルダパス>");
            //プログラム終了
            System.exit(1);
        }
        //csvのパスを宣言
        String csvPath = args[0];
        String folderPath = args[1];

        loadReplacementRules(csvPath);
        processFilesInFolder(folderPath);

    }

    private static void loadReplacementRules(String filePath) {
    	//テキストファイルを読み込む
        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(filePath), "UTF-8"))) {
            String line;
            while ((line = br.readLine()) != null) {
            	//コンマで区切る
                String[] parts = line.split(",");
                if (parts.length == 2) {
                	//マップに値を詰める
                    replacementMap.put(parts[0].trim(), parts[1].trim());
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void processFilesInFolder(String folderPath) {
        try {
        	//ファイルやディレクトリの一覧をサブディレクトリの中まで取得する
            Files.walk(Paths.get(folderPath))
                    .filter(Files::isRegularFile)
                    .forEach(path -> processFile(path.toFile()));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    //docx, xlsx, pptx によって異なるメソッドを呼び出す
    private static void processFile(File file) {
        String name = file.getName().toLowerCase();
        try {
        if (name.endsWith(".docx")) {
            processWordFile(file);
        } else if (name.endsWith(".xlsx")) {
            processExcelFile(file);
        } else if (name.endsWith(".pptx")) {
            processPowerPointFile(file);
        }
        //それ以外の場合はエラー処理。
        } catch (Exception e) {
        	System.err.println("Error processing file: " + file.getName());
            e.printStackTrace();
        }
    }

    // Word
    private static void processWordFile(File file) {
        try (FileInputStream fis = new FileInputStream(file);
        	//XWPFDocument:.docxワードファイルを生成、編集することができるAPI
             XWPFDocument doc = new XWPFDocument(fis)) {
            //段落を変換
        	for (XWPFParagraph para : doc.getParagraphs()) {
        	    String text = para.getText();
        	    //mapに従って文字列を置換する
        	    for (Map.Entry<String, String> entry : replacementMap.entrySet()) {
        	        text = text.replace(entry.getKey(), entry.getValue());
        	    }
                //Run＝段落中の書式が同じテキストのまとまり
        	    //Runが何軒含まれるかを数える
        	    int runCount = para.getRuns().size();
        	    //Runを削除する
        	    for (int i = runCount - 1; i >= 0; i--) {
        	    	//指定したインデックスのRunを段落から削除する
        	        para.removeRun(i);
        	    }

        	    // 新しいRunを追加。
        	    XWPFRun newRun = para.createRun();
        	    newRun.setText(text);
        	     // フォント設定
        	    newRun.setFontFamily(FONT_NAME); 
        	}

            try (FileOutputStream fos = new FileOutputStream(file)) {
                doc.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Excel
    private static void processExcelFile(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFFont font = workbook.createFont();
            font.setFontName(FONT_NAME);

            for (org.apache.poi.ss.usermodel.Sheet sheet : workbook) {
                for (Row row : sheet) {
                    for (org.apache.poi.ss.usermodel.Cell cell : row) {
                        if (cell.getCellType() == CellType.STRING) {
                            String text = cell.getStringCellValue();
                            for (Map.Entry<String, String> entry : replacementMap.entrySet()) {
                                text = text.replace(entry.getKey(), entry.getValue());
                            }
                            cell.setCellValue(text);
                        }
                        // スタイル適用
                        CellStyle oldStyle = cell.getCellStyle();
                        CellStyle newStyle = workbook.createCellStyle();
                        newStyle.cloneStyleFrom(oldStyle);
                        newStyle.setFont(font);
                        cell.setCellStyle(newStyle);
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

    // PowerPoint
    private static void processPowerPointFile(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             XMLSlideShow ppt = new XMLSlideShow(fis)) {

            for (XSLFSlide slide : ppt.getSlides()) {
                for (XSLFShape shape : slide.getShapes()) {
                	if (shape instanceof XSLFTextShape) {
                	    XSLFTextShape textShape = (XSLFTextShape) shape;
                	    updateTextShape(textShape);
                	}

                }
            }

            for (XSLFSlideMaster master : ppt.getSlideMasters()) {
                for (XSLFShape shape : master.getShapes()) {
                	if (shape instanceof XSLFTextShape) {
                	    XSLFTextShape textShape = (XSLFTextShape) shape;
                	    updateTextShape(textShape);
                	}
                }
            }

            modifyThemeFonts(ppt, FONT_NAME);

            try (FileOutputStream fos = new FileOutputStream(file)) {
                ppt.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    private static void updateTextShape(XSLFTextShape textShape) {
        // 元テキスト取得
        String originalText = textShape.getText();
        String replacedText = originalText;
        
        // 置換処理
        for (Map.Entry<String, String> entry : replacementMap.entrySet()) {
            replacedText = replacedText.replace(entry.getKey(), entry.getValue());
        }

        // 古い段落とランをすべて削除
        textShape.clearText();

        // 新しい段落とテキストランを追加
        XSLFTextParagraph para = textShape.addNewTextParagraph();
        XSLFTextRun run = para.addNewTextRun();
        run.setText(replacedText);
        run.setFontFamily(FONT_NAME);  // 必要であればフォント設定
    }
    //スライドマスターのフォントも変更する
	private static void modifyThemeFonts(XMLSlideShow ppt, String fontName) {
        for (XSLFSlideMaster master : ppt.getSlideMasters()) {
            XSLFTheme theme = master.getTheme();
            if (theme == null) continue;

            CTOfficeStyleSheet styleSheet = theme.getXmlObject();
            if (styleSheet.getThemeElements() != null &&
                styleSheet.getThemeElements().getFontScheme() != null) {

                CTFontScheme fontScheme = styleSheet.getThemeElements().getFontScheme();

                if (fontScheme.getMajorFont() != null && fontScheme.getMajorFont().getLatin() != null) {
                    fontScheme.getMajorFont().getLatin().setTypeface(fontName);
                }

                if (fontScheme.getMinorFont() != null && fontScheme.getMinorFont().getLatin() != null) {
                    fontScheme.getMinorFont().getLatin().setTypeface(fontName);
                }
            }
        }
    }


}
