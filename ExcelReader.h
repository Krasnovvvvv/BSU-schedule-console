#ifndef BSU_SCHEDULE_EXCELREADER_H
#define BSU_SCHEDULE_EXCELREADER_H

#include <xlsxio_read.h>
#include <string>
#include <unordered_map>
#include <vector>
#include <iostream>

class ExcelReader {

    const char* tablePath;
    xlsxioreader table;

    std::vector<const char*> findSheets() {
        xlsxioreadersheetlist sheetlist = xlsxioread_sheetlist_open(table);
        const char* sheetname;
        std::vector<const char*> sheets;

        while((sheetname = xlsxioread_sheetlist_next(sheetlist)) != nullptr) {
            sheets.emplace_back(sheetname);
        }
        xlsxioread_sheetlist_close(sheetlist);
        return sheets;
    }

public:

    std::vector<const char*> _sheets;

    std::unordered_map<const char*, std::pair<int,int>> readSheet(const char* sheet_name) {
        std::unordered_map<const char*, std::pair<int,int>> cellMap;
        int row = 0;
        xlsxioreadersheet sheet = xlsxioread_sheet_open(table,sheet_name, XLSXIOREAD_SKIP_ALL_EMPTY);
        while(xlsxioread_sheet_next_row(sheet)) {
            int col = 0;
            char* value;
            while((value=xlsxioread_sheet_next_cell(sheet))!= nullptr) {
                cellMap[value] = std::make_pair(row,col);
                xlsxioread_free(value);
                ++col;
            }
            ++row;
        }
        xlsxioread_sheet_close(sheet);
        return cellMap;
    }

    explicit ExcelReader(const char* path) : tablePath(path) {
        table = xlsxioread_open(tablePath);
        if(!table) {
            throw std::runtime_error("Invalid table path!");
        }
        findSheets();
    };

    explicit ExcelReader() : ExcelReader("C:/myTables") {};

    ~ExcelReader() {
        xlsxioread_close(table);
    }

};

#endif //BSU_SCHEDULE_EXCELREADER_H
