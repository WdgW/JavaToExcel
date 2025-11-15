package com.converter;

import java.io.OutputStream;
import com.github.javaparser.ParserConfiguration;
import com.github.javaparser.JavaParser;
import com.github.javaparser.ParseResult;
import com.github.javaparser.ast.CompilationUnit;
import com.github.javaparser.ast.body.FieldDeclaration;
import com.github.javaparser.ast.body.VariableDeclarator;
import com.github.javaparser.ast.expr.Expression;
import com.github.javaparser.ast.visitor.VoidVisitorAdapter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;
import java.util.stream.Collectors;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

@Command(
        name = "JavaToExcelConverter",
        mixinStandardHelpOptions = true,
        description = "Converts Java class fields to Excel with multiple sheets")
public class JavaToExcelConverter implements Callable<Integer> {

    private static final Logger logger = Logger.getLogger(JavaToExcelConverter.class.getName());

    @Option(
            names = {"-i", "--input-folder"},
            required = true,
            description = "Path to folder containing Java files")
    private String inputFolder;

    @Option(
            names = {"-o", "--output-root"},
            required = true,
            description = "输出根目录，会按输入目录层级生成子目录及最底层 Excel")
    private String outputRoot;

    // @Option(names = {"-o", "--output-file"}, required = true,
    //        description = "Output Excel file path")
    // private String outputFile;

    static class FieldInfo {
        String fieldName;
        String typeName;
        String defaultValue;
        String comment;

        FieldInfo(String fieldName, String typeName, String defaultValue, String comment) {
            this.fieldName = fieldName;
            this.typeName = typeName;
            this.defaultValue = defaultValue;
            this.comment = comment;
        }
    }

    public static void main(String[] args) {
        setupLogger();
        int exitCode = new CommandLine(new JavaToExcelConverter()).execute(args);
        System.exit(exitCode);
    }

    private static void setupLogger() {
        try {
            FileHandler fileHandler = new FileHandler("parsing_errors.log");
            fileHandler.setFormatter(new SimpleFormatter());
            logger.addHandler(fileHandler);
        } catch (IOException e) {
            logger.log(Level.SEVERE, "Failed to setup logger", e);
        }
    }
    private void buildExcel(List<Path> javaFiles, Path excelPath) {
    try (Workbook wb = new XSSFWorkbook()) {
        for (Path java : javaFiles) {
            List<FieldInfo> fields = parseJavaFile(java.toFile());
            if (fields.isEmpty()) continue;
            String sheetName = java.getFileName().toString()
                                   .replace(".java", "");
            if (sheetName.length() > 31) sheetName = sheetName.substring(0, 28) + "...";
            createSheet(wb, sheetName, fields);   // 你原来的方法
        }
        try (OutputStream out = Files.newOutputStream(excelPath)) {
            wb.write(out);
        }
        logger.info("Generated: " + excelPath);
    } catch (Exception e) {
        logger.log(Level.SEVERE, "write excel failed: " + excelPath, e);
    }
}


    @Override
    public Integer call() {
        //setupLogger();
        Path inRoot = Paths.get(inputFolder);
        Path outRoot = Paths.get(outputRoot);

        try {
            // 1. 遍历输入目录，按层创建输出目录
            Files.walk(inRoot)
                    .filter(Files::isDirectory)
                    .forEach(dir -> {
                        // 对应的输出目录
                        Path relative = inRoot.relativize(dir);
                        Path outDir = outRoot.resolve(relative);
                        try {
                            Files.createDirectories(outDir);
                        } catch (IOException e) {
                            logger.log(Level.SEVERE, "mkdir failed: " + outDir, e);
                        }

                        // 2. 如果当前目录已经是“最底层”（没有子目录），就生成 Excel
                        boolean hasSubDir;
                        try {
                            hasSubDir = Files.list(dir).anyMatch(Files::isDirectory);
                        } catch (IOException e) {
                            logger.log(Level.SEVERE, "list failed: " + dir, e);
                            return;
                        }
                        if (!hasSubDir) {
                            List<Path> javaFiles;
                            try {
                                javaFiles = Files.list(dir)
                                        .filter(p -> p.toString().endsWith(".java"))
                                        .collect(Collectors.toList());
                            } catch (IOException e) {
                                logger.log(Level.SEVERE, "list java files failed: " + dir, e);
                                return;
                            }
                            if (!javaFiles.isEmpty()) {
                                String excelName = dir.getFileName() + ".xlsx";
                                Path excelPath = outDir.resolve(excelName);
                                buildExcel(javaFiles, excelPath);
                            }
                        }
                    });
            logger.info("All done. Check output under: " + outRoot.toAbsolutePath());
            return 0;
        } catch (Exception e) {
            logger.log(Level.SEVERE, "Fatal error", e);
            return 1;
        }
    }

    private List<FieldInfo> parseJavaFile(File file) throws IOException {
        List<FieldInfo> fields = new ArrayList<>();
        ParserConfiguration config = new ParserConfiguration()
         .setLanguageLevel(ParserConfiguration.LanguageLevel.JAVA_21);
        JavaParser parser = new JavaParser(config);

        try (FileInputStream fis = new FileInputStream(file)) {
            ParseResult<CompilationUnit> result = parser.parse(fis);

            if (result.isSuccessful() && result.getResult().isPresent()) {
                CompilationUnit cu = result.getResult().get();

                new VoidVisitorAdapter<Void>() {
                    @Override
                    public void visit(FieldDeclaration field, Void arg) {
                        super.visit(field, arg);
                        String typeName = field.getElementType().asString();
                        String comment = field.getJavadoc().map(jd -> jd.getDescription().toText()).orElse(null);

                        for (VariableDeclarator var : field.getVariables()) {
                            String defaultValue = var.getInitializer()
                                    .map(Expression::toString)
                                    .orElse(null);

                            fields.add(new FieldInfo(
                            var.getNameAsString(),
                            typeName,
                            defaultValue,
                            comment
                            ));
                        }
                    }
                }.visit(cu, null);
            } else {
                logger.severe("Failed to parse file: " + file.getAbsolutePath());
            }
        }

        return fields;
    }

    private void createSheet(Workbook workbook, String fileName, List<FieldInfo> fields) {
        String sheetName = fileName.replace(".java", "");
        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 28) + "...";
        }

        Sheet sheet = workbook.createSheet(sheetName);
        Row headerRow = sheet.createRow(0);

        String[] headers = {"字段名", "类型", "默认值", "注释"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        int rowNum = 1;
        for (FieldInfo field : fields) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(field.fieldName);
            row.createCell(1).setCellValue(field.typeName);

            if (field.defaultValue != null) {
                row.createCell(2).setCellValue(field.defaultValue);
            }

            if (field.comment != null) {
                row.createCell(3).setCellValue(field.comment);
            }
        }

        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private static final class Counter {
        int value;

        void inc() {
            value++;
        }

        int get() {
            return value;
        }
    }
                       }
