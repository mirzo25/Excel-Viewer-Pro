package com.excelviewer.pro;

import android.app.Activity;
import android.os.Bundle;
import android.content.Intent;
import android.net.Uri;
import android.graphics.Color;
import android.widget.*;

import org.apache.poi.ss.usermodel.*;

import java.io.InputStream;
import java.util.Locale;

public class MainActivity extends Activity {

    private Workbook workbook;
    private LinearLayout tabs;
    private TextView result;
    private EditText find;
    private DataFormatter formatter;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        formatter = new DataFormatter(Locale.getDefault());

        LinearLayout root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);
        root.setPadding(8, 8, 8, 8);
        root.setBackgroundColor(Color.WHITE);

        TextView title = new TextView(this);
        title.setText("EXCEL VIEWER PRO");
        title.setTextSize(22);
        title.setTextColor(Color.BLACK);
        title.setGravity(17);
        title.setPadding(5, 15, 5, 15);
        root.addView(title);

        Button open = new Button(this);
        open.setText("ОТКРЫТЬ EXCEL-ФАЙЛ");
        root.addView(open);
        open.setOnClickListener(v -> openFile());

        LinearLayout searchLayout = new LinearLayout(this);
        searchLayout.setOrientation(LinearLayout.HORIZONTAL);

        find = new EditText(this);
        find.setHint("Название или код товара");
        find.setSingleLine(true);

        searchLayout.addView(find,
                new LinearLayout.LayoutParams(0, -2, 1));

        Button searchButton = new Button(this);
        searchButton.setText("НАЙТИ");
        searchLayout.addView(searchButton);
        root.addView(searchLayout);

        searchButton.setOnClickListener(v -> findText());

        HorizontalScrollView horizontal =
                new HorizontalScrollView(this);

        tabs = new LinearLayout(this);
        tabs.setOrientation(LinearLayout.HORIZONTAL);
        horizontal.addView(tabs);
        root.addView(horizontal);

        ScrollView scroll = new ScrollView(this);
        result = new TextView(this);
        result.setTextColor(Color.BLACK);
        result.setTextSize(14);
        result.setPadding(8, 12, 8, 12);
        scroll.addView(result);

        root.addView(scroll,
                new LinearLayout.LayoutParams(-1, 0, 1));

        result.setText("Выберите Excel-файл XLS или XLSX.");
        setContentView(root);
    }private void openFile() {
        Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        intent.setType("*/*");
        intent.addCategory(Intent.CATEGORY_OPENABLE);

        intent.putExtra(Intent.EXTRA_MIME_TYPES, new String[]{
            "application/vnd.ms-excel",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        });

        startActivityForResult(intent, 100);
    }

    @Override
    protected void onActivityResult(
            int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        if (requestCode == 100 && resultCode == RESULT_OK
                && data != null && data.getData() != null) {
            loadFile(data.getData());
        }
    }

    private void loadFile(Uri uri) {
        Workbook newWorkbook = null;

        try {
            InputStream input =
                    getContentResolver().openInputStream(uri);

            if (input == null) {
                throw new Exception("Не удалось открыть файл");
            }

            try {
                newWorkbook = WorkbookFactory.create(input);
            } finally {
                input.close();
            }

            if (newWorkbook.getNumberOfSheets() == 0) {
                newWorkbook.close();
                throw new Exception("В файле нет листов");
            }

            if (workbook != null) {
                workbook.close();
            }

            workbook = newWorkbook;
            showTabs();
            showSheet(0);

        } catch (Exception e) {
            try {
                if (newWorkbook != null && newWorkbook != workbook) {
                    newWorkbook.close();
                }
            } catch (Exception ignored) {
            }

            result.setText("Ошибка: " + e.getMessage());
        }
    }

    private void showTabs() {
        tabs.removeAllViews();

        if (workbook == null) return;

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            final int number = i;

            Button tab = new Button(this);
            tab.setText(workbook.getSheetName(i));
            tab.setOnClickListener(v -> showSheet(number));

            tabs.addView(tab);
        }
    }private void showSheet(int index) {
        if (workbook == null) return;

        try {
            Sheet sheet = workbook.getSheetAt(index);
            StringBuilder output = new StringBuilder();

            for (Row row : sheet) {
                boolean hasData = false;

                for (Cell cell : row) {
                    String value =
                            formatter.formatCellValue(cell);

                    output.append(value).append(" | ");

                    if (!value.trim().isEmpty()) {
                        hasData = true;
                    }
                }

                if (hasData) {
                    output.append("\n");
                }
            }

            if (output.length() == 0) {
                result.setText("Этот лист пустой.");
            } else {
                result.setText(output.toString());
            }

        } catch (Exception e) {
            result.setText("Ошибка чтения листа: " + e.getMessage());
        }
    }

    private void findText() {
        if (workbook == null) {
            result.setText("Сначала откройте Excel-файл.");
            return;
        }

        String query = find.getText().toString()
                .trim().toLowerCase(Locale.ROOT);

        if (query.isEmpty()) {
            result.setText("Введите название или код товара.");
            return;
        }

        StringBuilder found = new StringBuilder();

        try {
            for (int s = 0;
                    s < workbook.getNumberOfSheets(); s++) {

                Sheet sheet = workbook.getSheetAt(s);

                for (Row row : sheet) {
                    StringBuilder rowText = new StringBuilder();
                    boolean matched = false;

                    for (Cell cell : row) {
                        String value =
                                formatter.formatCellValue(cell);

                        rowText.append(value).append(" | ");

                        if (value.toLowerCase(Locale.ROOT)
                                .contains(query)) {
                            matched = true;
                        }
                    }

                    if (matched) {
                        found.append("Лист: ")
                                .append(sheet.getSheetName())
                                .append("\n");

                        found.append("Строка: ")
                                .append(row.getRowNum() + 1)
                                .append("\n");

                        found.append(rowText)
                                .append("\n\n");
                    }
                }
            }

            if (found.length() == 0) {
                result.setText("Ничего не найдено.");
            } else {
                result.setText(found.toString());
            }

        } catch (Exception e) {
            result.setText("Ошибка поиска: " + e.getMessage());
        }
    }

    @Override
    protected void onDestroy() {
        try {
            if (workbook != null) {
                workbook.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        super.onDestroy();
    }
}
