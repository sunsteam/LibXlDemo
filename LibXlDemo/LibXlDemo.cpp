#include <complex>
#include <iostream>
#include <map>
#include <vector>

#include "includes/libxl.h"

#pragma comment(lib,"libs/libxl.lib")

using namespace libxl;

enum Formats { NONE, ALIGN, PERCENT, DECIMAL, TITLE, SCIENTIFIC };

std::map<Formats, IFormatT<wchar_t>*> formatMap;

auto* initBook(const wchar_t *suffix) {
    const auto book = wcscmp(suffix, L"xlsx") == 0 ? xlCreateXMLBook() : xlCreateBook();
    book->setDefaultFont(L"Consolas", 10);
    //1
    auto *alignFormat = book->addFormat();
    alignFormat->setAlignH(ALIGNH_CENTER);
    alignFormat->setAlignV(ALIGNV_CENTER);
    //2
    auto *percentFormat = book->addFormat(alignFormat);
    percentFormat->setNumFormat(NUMFORMAT_PERCENT_D2);
    //3
    auto *decimalFormat = book->addFormat(alignFormat);
    decimalFormat->setNumFormat(book->addCustomNumFormat(L"0.000"));
    //4
    auto *titleFormat = book->addFormat(alignFormat);
    auto *boldFont = book->addFont();
    boldFont->setBold(true);
    titleFormat->setFont(boldFont);
    //5
    auto *scientificFormat = book->addFormat(alignFormat);
    scientificFormat->setNumFormat(NUMFORMAT_SCIENTIFIC_D2);

    formatMap[ALIGN] = alignFormat;
    formatMap[PERCENT] = percentFormat;
    formatMap[DECIMAL] = decimalFormat;
    formatMap[TITLE] = titleFormat;
    formatMap[SCIENTIFIC] = scientificFormat;
    return book;
}

auto writeExcel(const int totalCount, int passCount, int failCount, double average, double yield,
                const std::vector<double> &dataList, const std::wstring &&filePath) -> void {
    const auto suffix = filePath.substr(filePath.rfind(L'.') + 1);
    auto *book = initBook(suffix.c_str());
    auto *sheet = book->addSheet(L"Test Sheet");
    int row = 0;
    constexpr int headCol = 0;
    constexpr int valCol = 1;

#if 0
    // if use map to save pointer, no problem.
    sheet->setCol(0, 4, 12, formatMap[ALIGN]);
    const auto titleFormat = formatMap[TITLE];
    sheet->writeStr(row++, headCol, L"Test LibXl Functions", titleFormat);
    sheet->writeStr(row, headCol, L"totalCount", titleFormat);
    sheet->writeNum(row++, valCol, totalCount);
    sheet->writeStr(row, headCol, L"passCount", titleFormat);
    sheet->writeNum(row++, valCol, passCount);
    sheet->writeStr(row, headCol, L"failCount", titleFormat);
    sheet->writeNum(row++, valCol, failCount);
    sheet->writeStr(row, headCol, L"average", titleFormat);
    sheet->writeNum(row++, valCol, average, formatMap[DECIMAL]);
    sheet->writeStr(row, headCol, L"yield", titleFormat);
    sheet->writeNum(row++, valCol, yield, formatMap[PERCENT]);
    row++;

    sheet->writeStr(row, 0, L"Data", titleFormat);
    sheet->writeStr(row++, 1, L"Scientific", titleFormat);
    for (const auto data : dataList) {
        sheet->writeNum(row, 0, data, formatMap[DECIMAL]);
        sheet->writeNum(row++, 1, data, formatMap[SCIENTIFIC]);
    }
#else
    // use book->format() function, obtain wrong format pointer when use xlCreateBook(), can't be work
    sheet->setCol(0, 4, 12, book->format(ALIGN));
    const auto titleFormat = book->format(TITLE);
    sheet->writeStr(row++, headCol, L"Test LibXl Functions", titleFormat);
    sheet->writeStr(row, headCol, L"totalCount", titleFormat);
    sheet->writeNum(row++, valCol, totalCount);
    sheet->writeStr(row, headCol, L"passCount", titleFormat);
    sheet->writeNum(row++, valCol, passCount);
    sheet->writeStr(row, headCol, L"failCount", titleFormat);
    sheet->writeNum(row++, valCol, failCount);
    sheet->writeStr(row, headCol, L"average", titleFormat);
    sheet->writeNum(row++, valCol, average, book->format(DECIMAL));
    sheet->writeStr(row, headCol, L"yield", titleFormat);
    sheet->writeNum(row++, valCol, yield, book->format(PERCENT));
    row++;

    sheet->writeStr(row, 0, L"Data", titleFormat);
    sheet->writeStr(row++, 1, L"Scientific", titleFormat);
    for (const auto data : dataList) {
        sheet->writeNum(row, 0, data, book->format(DECIMAL));
        sheet->writeNum(row++, 1, data, book->format(SCIENTIFIC));
    }
#endif

    book->save(filePath.c_str());
    book->release();
}

int main() {
    std::cout << "Hello LibXl!\n";

    // calculate data;
    constexpr int totalCount = 100;
    int passCount = 0;
    int failCount = 0;
    double average = 0;
    double yield = 0;
    std::vector<double> dataList;
    srand(time(nullptr));
    for (int i = 0; i < totalCount; ++i) {
        const int x = rand() % 100 + 1;
        const int pow = x % 5;
        double data = static_cast<double>(x) * std::pow(10, -1 * pow);
        dataList.push_back(data);
        if (data > 1) {
            passCount++;
        } else {
            failCount++;
        }
        average += data;
        if (i == totalCount - 1) {
            average /= totalCount;
            yield = static_cast<double>(passCount) / totalCount;
        }
    }

    // write excel
    writeExcel(totalCount, passCount, failCount, average, yield, dataList, L"test.xlsx");
    writeExcel(totalCount, passCount, failCount, average, yield, dataList, L"test.xls");
}
