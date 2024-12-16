import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.chart.ChartUtils;

import java.io.*;
import java.util.*;

public class StudentAnalysis {

    public static void main(String[] args) {
        String inputFilePath = "C:/kr/abc.xlsx";  // Путь к входному файлу
        String outputFilePath = "result.xlsx";  // Путь к выходному файлу

        try {
            // Чтение входного Excel файла
            List<Student> students = readExcelFile(inputFilePath);

            // Анализ данных
            Statistics stats = analyzeGrades(students);

            // Запись результатов в новый Excel файл
            writeResultsToExcel(outputFilePath, stats);
            System.out.println("Результаты успешно записаны в файл: " + outputFilePath);

            // Генерация графика
            generateChart(stats);

        } catch (Exception e) {
            System.out.println("Произошла ошибка: " + e.getMessage());
        }
    }

    // Чтение Excel файла и извлечение данных студентов
    public static List<Student> readExcelFile(String filePath) throws IOException {
        List<Student> students = new ArrayList<>();
        FileInputStream file = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue; // Пропуск заголовка
            String name = row.getCell(0).getStringCellValue();
            int grade = (int) row.getCell(1).getNumericCellValue();
            students.add(new Student(name, grade));
        }

        workbook.close();
        return students;
    }

    // Анализ оценок
    public static Statistics analyzeGrades(List<Student> students) {
        Statistics stats = new Statistics();

        for (Student student : students) {
            if (student.getGrade() == 5) stats.getExcellent().add(student);
            else if (student.getGrade() == 4) stats.getGood().add(student);
            else if (student.getGrade() == 3) stats.getSatisfactory().add(student);
            else stats.getFailed().add(student);
        }

        stats.setAverageGrade(students.stream().mapToInt(Student::getGrade).average().orElse(0));
        return stats;
    }

    // Запись результатов в новый Excel файл
    public static void writeResultsToExcel(String filePath, Statistics stats) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Результаты");

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ФИО");
        headerRow.createCell(1).setCellValue("Оценка");

        int rowNum = 1;

        // Заполнение статистики
        Row resultRow = sheet.createRow(rowNum++);
        resultRow.createCell(0).setCellValue("Средний балл");
        resultRow.createCell(1).setCellValue(stats.getAverageGrade());

        // Заполнение данных по группам
        createGroupRow(sheet, rowNum++, "Отличники", stats.getExcellent());
        createGroupRow(sheet, rowNum++, "Хорошисты", stats.getGood());
        createGroupRow(sheet, rowNum++, "Троешники", stats.getSatisfactory());
        createGroupRow(sheet, rowNum++, "Не допущен", stats.getFailed());

        // Сохранение в файл
        FileOutputStream fileOut = new FileOutputStream(filePath);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }

    // Создание строки для группы студентов
    public static void createGroupRow(Sheet sheet, int rowNum, String label, List<Student> students) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(label);
        int colNum = 1;
        for (Student student : students) {
            row.createCell(colNum++).setCellValue(student.getName());
        }
    }

    // Генерация графика
    public static void generateChart(Statistics stats) throws IOException {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        dataset.addValue(stats.getExcellent().size(), "Оценки", "5");
        dataset.addValue(stats.getGood().size(), "Оценки", "4");
        dataset.addValue(stats.getSatisfactory().size(), "Оценки", "3");
        dataset.addValue(stats.getFailed().size(), "Оценки", "Не допущен");

        JFreeChart chart = ChartFactory.createBarChart(
                "Распределение оценок",
                "Оценка", "Количество студентов",
                dataset, PlotOrientation.VERTICAL, true, true, false);

        File chartFile = new File("chart.png");
        ChartUtils.saveChartAsPNG(chartFile, chart, 600, 400);
        System.out.println("График сохранен в файл: " + chartFile.getAbsolutePath());
    }

    // Классы данных
    public static class Student {
        private String name;
        private int grade;

        public Student(String name, int grade) {
            this.name = name;
            this.grade = grade;
        }

        public String getName() {
            return name;
        }

        public int getGrade() {
            return grade;
        }
    }

    public static class Statistics {
        private List<Student> excellent = new ArrayList<>();
        private List<Student> good = new ArrayList<>();
        private List<Student> satisfactory = new ArrayList<>();
        private List<Student> failed = new ArrayList<>();
        private double averageGrade;

        public List<Student> getExcellent() {
            return excellent;
        }

        public List<Student> getGood() {
            return good;
        }

        public List<Student> getSatisfactory() {
            return satisfactory;
        }

        public List<Student> getFailed() {
            return failed;
        }

        public double getAverageGrade() {
            return averageGrade;
        }

        public void setAverageGrade(double averageGrade) {
            this.averageGrade = averageGrade;
        }
    }
}