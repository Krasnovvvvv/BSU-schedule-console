#ifndef BSU_SCHEDULE_EXCELREADER_H
#define BSU_SCHEDULE_EXCELREADER_H

#include <xlsxio_read.h>
#include <string>
#include <unordered_map>
#include <vector>
#include <iostream>

class ExcelReader {
    const char* tablePath;
    std::unordered_map<const char*, std::pair<int,int>> cellMap {};
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

    explicit ExcelReader(const char* path) : tablePath(path) {
        table = xlsxioread_open(tablePath);
        if(!table) {
            throw std::runtime_error("Invalid table path!");
        }
    };

    explicit ExcelReader() : ExcelReader("C:/myTables") {};

    ~ExcelReader() {
        xlsxioread_close(table);
    }


};

#endif //BSU_SCHEDULE_EXCELREADER_H
