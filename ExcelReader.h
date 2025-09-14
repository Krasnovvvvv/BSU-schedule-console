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
