import com.github.javaparser.JavaParser;
import com.github.javaparser.ParseResult;
import com.github.javaparser.ast.CompilationUnit;
import com.github.javaparser.ast.body.FieldDeclaration;
import com.github.javaparser.ast.body.VariableDeclarator;
import com.github.javaparser.ast.expr.Expression;
import com.github.javaparser.ast.visitor.VoidVisitorAdapter;
import com.github.javaparser.utils.ProjectRoot;
import com.github.javaparser.utils.SourceRoot;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;

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

@Command(name = "JavaToExcelConverter", mixinStandardHelpOptions = true, 
         description = "Converts Java class fields to Excel with multiple sheets")
public class JavaToExcelConverter implements Callable<Integer> {

    private static final Logger logger = Logger.getLogger(JavaToExcelConverter.class.getName());

    @Option(names = {"-i", "--input-folder"}, required = true, 
            description = "Path to folder containing Java files")
    private String inputFolder;

    @Option(names = {"-o", "--output-file"}, required = true, 
            description = "Output Excel file path")
    private String outputFile;

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

    @Override
    public Integer call() {
        try (Workbook workbook = new XSSFWorkbook()) {
            int successCount = 0;
            int errorCount = 0;

            Files.walk(Paths.get(inputFolder))
                    .filter(Files::isRegularFile)
                    .filter(path -> path.toString().endsWith(".java"))
                    .forEach(path -> {
                        try {
                            List<FieldInfo> fields = parseJavaFile(path.toFile());
                            if (!fields.isEmpty()) {
                                createSheet(workbook, path.getFileName().toString(), fields);
                                successCount++;
                            }
                        } catch (Exception e) {
                            logger.log(Level.SEVERE, "Error processing file: " + path, e);
                            errorCount++;
                        }
                    });

            logger.info("Processing completed: " + successCount + " files succeeded, " + errorCount + " files failed");
            saveWorkbook(workbook);
            System.out.println("Excel file generated: " + outputFile);
            System.out.println("Error log saved to: parsing_errors.log");
            return 0;
        } catch (Exception e) {
            logger.log(Level.SEVERE, "Fatal error during conversion", e);
            return 1;
        }
    }

    private List<FieldInfo> parseJavaFile(File file) throws IOException {
        List<FieldInfo> fields = new ArrayList<>();
        JavaParser parser = new JavaParser();
        
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

    private void saveWorkbook(Workbook workbook) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            workbook.write(fos);
        }
    }
}