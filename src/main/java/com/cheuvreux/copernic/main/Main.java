package com.cheuvreux.copernic.main;

import com.cheuvreux.copernic.cli.CLIParameters;
import me.tongfei.progressbar.DelegatingProgressBarConsumer;
import me.tongfei.progressbar.ProgressBar;
import me.tongfei.progressbar.ProgressBarBuilder;
import me.tongfei.progressbar.ProgressBarStyle;
import org.apache.commons.cli.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class Main {

    private static final Logger logger = LogManager.getLogger(Main.class);
    private Integer count = 0;
    private static String wordFile = null;
    private static String excelFile = null;
    private static Integer sheetIndex = null;
    private static Integer columnSourceIndex = null;
    private static Integer columnDestinationIndex = null;

    public static void main(String[] args) {
        try {
            final Options firstOptions = CLIParameters.configFirstParameters();
            final Options options = CLIParameters.configParameters(firstOptions);
            final CommandLineParser parser = new DefaultParser();
            final CommandLine firstLine = parser.parse(firstOptions, args, true);


            CLIParameters.helpMode(options, firstLine);

            final CommandLine line = parser.parse(options, args);
            excelFile = line.getOptionValue("excel");
            wordFile = line.getOptionValue("word");
            sheetIndex = line.getOptionValue("sheet-index")!=null? Integer.parseInt(line.getOptionValue("sheet-index")): 0;
            columnSourceIndex = line.getOptionValue("column-source-index") !=null ?Integer.parseInt(line.getOptionValue("column-source-index")): 2;
            columnDestinationIndex = line.getOptionValue("column-destination-index") !=null ? Integer.parseInt(line.getOptionValue("column-destination-index")): 3;

            logger.info("Launch program with values Excel File : {}, Word Input File : {}",
                    excelFile, wordFile);
            logger.info("======> Program Start <=======");

            new Main().extractFromExcelAndReplaceInWord();
            logger.info("======> Program Finished with success <=======");
        } catch (IOException | ParseException e) {
            logger.error("the program has exited with error : {}", e.getMessage());
        }
    }

    private void extractFromExcelAndReplaceInWord() throws IOException {
        changeValuesInWord(getExcelValues());
    }

    private void changeValuesInWord(Map<String, String> excelValues) throws IOException {
        logger.info(" start replacing source value with destination value in Word File");
        InputStream docFile = new FileInputStream(wordFile);
        try (XWPFDocument doc = new XWPFDocument(docFile)) {
            List<XWPFRun> runsCollect = doc.getParagraphs().stream()
                    .map(XWPFParagraph::getRuns).flatMap(List::stream)
                    .filter(Objects::nonNull).collect(Collectors.toList());
            List<XWPFRun> tableStream = doc.getTables().stream()
                    .map(XWPFTable::getRows).flatMap(List::stream)
                    .map(XWPFTableRow::getTableCells).flatMap(List::stream)
                    .map(XWPFTableCell::getParagraphs).flatMap(List::stream)
                    .map(XWPFParagraph::getRuns).flatMap(List::stream).collect(Collectors.toList());
            ProgressBarBuilder replacing = new ProgressBarBuilder()
                    .setTaskName("Replacing").setStyle(ProgressBarStyle.ASCII);
            ProgressBar.wrap(excelValues.entrySet(), replacing).forEach(entry -> {
                runsCollect.parallelStream().forEach(r -> replaceText(entry, r));
                tableStream.parallelStream().forEach(r -> replaceText(entry, r));
            });
            try (FileOutputStream out = new FileOutputStream(wordFile)) {
                doc.write(out);
            }
        }
        logger.info(" Value change in Word File with Success : number of word changed : {}", count);
    }

    private Map<String, String> getExcelValues() throws IOException {
        logger.info("start extract source and destination values from excel File");
        try (InputStream excelFileInputStream = new FileInputStream(Main.excelFile)) {
            try (Workbook sheets = WorkbookFactory.create(excelFileInputStream)) {
                Sheet sheetAt = sheets.getSheetAt(sheetIndex);
                Map<String, String> excelValues = new HashMap<>();
                if (sheetAt != null) {
                    Stream<Row> rowStream = StreamSupport.stream(
                            Spliterators.spliteratorUnknownSize(sheetAt.rowIterator(), Spliterator.ORDERED),
                            false);
                    rowStream.parallel().skip(1).forEach(row -> {
                        if (row != null && row.getCell(columnSourceIndex) != null && row.getCell(columnDestinationIndex) != null) {
                            String source = row.getCell(columnSourceIndex).getStringCellValue();
                            String destination = row.getCell(columnDestinationIndex).getStringCellValue();
                            excelValues.put(source, destination);
                        }
                    });
                }
                logger.info("excel values extracted with success, total lines : {}",excelValues.size());
                return excelValues;
            }
        }
    }

    private void replaceText(Map.Entry<String, String> entry, XWPFRun r) {
        String text = r.getText(0);
        if (text != null && text.toLowerCase().contains(entry.getKey().toLowerCase())) {
            text = text.replace(entry.getKey(), entry.getValue());
            logger.info("found : {}", entry.getKey());
            logger.info("replace with  : {}", entry.getValue());
            count++;
            r.setText(text, 0);
        }
    }
}
