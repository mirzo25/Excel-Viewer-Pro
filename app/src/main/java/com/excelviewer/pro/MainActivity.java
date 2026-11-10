for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            final int index = i;

            Button button = new Button(this);
            button.setText(workbook.getSheetName(i));

            button.setOnClickListener(v -> showSheet(index));

            sheetBar.addView(button);
        }
    }

    private void showSheet(int index) {
        if (workbook == null) return;

        currentSheet = index;
        Sheet sheet = workbook.getSheetAt(index);

        StringBuilder result = new StringBuilder();

        result.append("Лист: ")
                .append(sheet.getSheetName())
                .append("\n\n");

        for (Row row : sheet) {
            for (Cell cell : row) {
                result.append(formatter.formatCellValue(cell))
                        .append("    ");
            }
            result.append("\n");
        }

        content.setText(result.toString());
    }

    private void searchExcel() {
        if (workbook == null) {
            content.setText("Сначала откройте Excel-файл.");
            return;
        }

        String query = searchBox.getText().toString()
                .trim().toLowerCase(Locale.ROOT);

        if (query.isEmpty()) {
            content.setText("Введите название товара или код.");
            return;
        }

        StringBuilder result = new StringBuilder();
        int count = 0;

        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            Sheet sheet = workbook.getSheetAt(s);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    String value = formatter.formatCellValue(cell);

                    if (value.toLowerCase(Locale.ROOT).contains(query)) {
                        result.append("Лист: ")
                                .append(sheet.getSheetName())
                                .append("\nСтрока: ")
                                .append(row.getRowNum() + 1)
                                .append("\n")
                                .append(value)
                                .append("\n\n");

                        count++;

                        if (count >= 500) break;
                    }
                }
                if (count >= 500) break;
            }
            if (count >= 500) break;
        }

        if (count == 0) {
            result.append("Ничего не найдено.");
        } else {
            result.insert(0, "Найдено совпадений: " + count + "\n\n");
        }

        content.setText(result.toString());
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();

        try {
            if (workbook != null) {
                workbook.close();
            }
        } catch (Exception ignored) {
        }
    }
}
