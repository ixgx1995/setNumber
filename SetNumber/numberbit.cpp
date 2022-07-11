#include "excel.h"
#include "numberbit_p.h"
#include "numberbit.h"
#include <QtMath>
#include <QDir>
#include <QString>
#include <QDebug>

NumberBit::NumberBit(QObject *parent)
    : QObject(parent), d_ptr(new NumberBitPrivate), actExcel(nullptr)
{
}

NumberBit::~NumberBit()
{
    closeDocument();
}

QList<QString> NumberBit::documentNames() const
{
    return d_ptr->documentNames;
}

void NumberBit::setDocumentNames(QList<QString> value)
{
    if(d_ptr->documentNames == value)
        return;
    d_ptr->documentNames = value;
    emit documentNamesChanged(value);
}

void NumberBit::addDocumentName(QString value)
{
    if(d_ptr->documentNames.contains(value))
        return;
    d_ptr->documentNames.append(value);
    emit documentNamesChanged(d_ptr->documentNames);
}

void NumberBit::delDocumentName(QString value)
{
    if(documentNames().contains(value)) {
        d_ptr->documentNames.removeOne(value);
        emit documentNamesChanged(d_ptr->documentNames);
    }
}

void NumberBit::cleDocumentName()
{
    d_ptr->documentNames.clear();
    emit documentNamesChanged(d_ptr->documentNames);
}

QList<int> NumberBit::documentIndexs() const
{
    return d_ptr->documentIndexs;
}

void NumberBit::setdocumentIndexs(QList<int> value)
{
    if(d_ptr->documentIndexs == value)
        return;
    d_ptr->documentIndexs = value;
    emit documentIndexsChanged(value);
}

QList<QString> NumberBit::bookNames() const
{
    return d_ptr->bookNames;
}

void NumberBit::setBookNames(QList<QString> value)
{
    if(d_ptr->bookNames == value)
        return;
    d_ptr->bookNames = value;
    emit bookNamesChanged(value);
}

void NumberBit::addBookName(QString value)
{
    if(d_ptr->bookNames.contains(value))
        return;
    d_ptr->bookNames.append(value);
    emit bookNamesChanged(d_ptr->bookNames);
}

void NumberBit::delBookName(QString value)
{
    if(d_ptr->bookNames.contains(value)) {
        d_ptr->bookNames.removeOne(value);
        emit bookNamesChanged(d_ptr->bookNames);
    }
}

void NumberBit::cleBookName()
{
    d_ptr->bookNames.clear();
    emit bookNamesChanged(d_ptr->bookNames);
}

QList<int> NumberBit::workBookIndexs() const
{
    return d_ptr->workBookIndexs;
}

void NumberBit::setWorkBookIndexs(QList<int> value)
{
    if(d_ptr->workBookIndexs == value)
        return;
    d_ptr->workBookIndexs = value;
    emit workBookIndexsChanged(value);
}

QList<QString> NumberBit::sheetNames() const
{
    return d_ptr->sheetNames;
}

void NumberBit::setSheetNames(QList<QString> value)
{
    if(d_ptr->sheetNames == value)
        return;
    d_ptr->sheetNames = value;
    emit sheetNamesChanged(value);
}

void NumberBit::addSheetName(QString value)
{
    if(d_ptr->sheetNames.contains(value))
        return;
    d_ptr->sheetNames.append(value);
    emit sheetNamesChanged(d_ptr->sheetNames);
}

void NumberBit::delSheetName(QString value)
{
    if(d_ptr->sheetNames.contains(value)) {
        d_ptr->sheetNames.append(value);
        emit sheetNamesChanged(d_ptr->sheetNames);
    }
}

void NumberBit::cleSheetName()
{
    d_ptr->sheetNames.clear();
    emit sheetNamesChanged(d_ptr->sheetNames);
}

QList<int> NumberBit::workSheetIndexs() const
{
    return d_ptr->workSheetIndexs;
}

void NumberBit::setWorkSheetIndexs(QList<int> value)
{
    if(d_ptr->workSheetIndexs == value)
        return;
    d_ptr->workSheetIndexs = value;
    emit workSheetIndexsChanged(value);
}

QList<QString> NumberBit::colNames() const
{
    return d_ptr->colNames;
}

void NumberBit::setColNames(QList<QString> value)
{
    if(d_ptr->colNames == value)
        return;
    d_ptr->colNames = value;
    emit colNamesChanged(value);
}

void NumberBit::addColName(QString value)
{
    if(d_ptr->colNames.contains(value))
        return;
    d_ptr->colNames.append(value);
    emit colNamesChanged(d_ptr->colNames);
}

void NumberBit::delColName(QString value)
{
    if(d_ptr->colNames.contains(value)) {
        d_ptr->colNames.append(value);
        emit colNamesChanged(d_ptr->colNames);
    }
}

void NumberBit::cleColName()
{
    d_ptr->colNames.clear();
    emit colNamesChanged(d_ptr->colNames);
}

QList<int> NumberBit::colIndexs() const
{
    return d_ptr->colIndexs;
}

void NumberBit::setColIndexs(QList<int> value)
{
    if(d_ptr->colIndexs == value)
        return;
    d_ptr->colIndexs = value;

}

int NumberBit::numberBit() const
{
    return d_ptr->numberBit;
}

void NumberBit::setNumberBit(int value)
{
    if(d_ptr->numberBit == value)
        return;
    d_ptr->numberBit = value;
    emit numberBitChanged(value);
}

void NumberBit::closeDocument()
{
    auto keys = excels.keys();
    for(int i = 0; i < keys.length(); i++) {
        delete excels[keys[i]];
        excels[keys[i]] = nullptr;
    }
    excels.clear();
}

void NumberBit::openDocument(QString fpath)
{
    QString kName = fpath.split(".", QString::SkipEmptyParts).last();
    if((kName.contains("xlsx") || kName.contains("xls")) == false) {
        return;
    }

    QString name = fpath.split("/", QString::SkipEmptyParts).last();
    if(documentNames().contains(name)) {
        return;
    }
    addDocumentName(name);
    QExcel* excel = new QExcel(fpath);
    excels.insert(fpath, excel);
}

void NumberBit::selDocument(QString documentName)
{
    // 全选或多选
    if(documentName.isEmpty()) {
        cleSheetName();
        selWorkSheet("");
        actExcel = nullptr;
        return;
    }

    QString key;
    auto keys = excels.keys();
    for(int n = 0; n < keys.length(); n++) {
        if(keys[n].contains(documentName))
            key = keys[n];
    }

    auto excel = excels[key];
    for(int i = 1; i <= excel->getSheetsCount(); i++) {
        addSheetName(excel->getSheetName(i));
    }
    actExcel = excel;
}

void NumberBit::selWorkBook(QString bookName)
{
    if(bookName.isEmpty()) {
        cleSheetName();
        selWorkSheet("");
        return;
    }

    auto excel = actExcel;
    excel->selectWorkBook(bookName);
    for(int i = 1; i <= excel->getSheetsCount(); i++) {
        addSheetName(excel->getSheetName(i));
    }
}

void NumberBit::selWorkSheet(QString sheetName)
{
    if(sheetName.isEmpty()) {
        cleColName();
        return;
    }

    auto excel = actExcel;
    excel->selectSheet(sheetName);
    for(int i = excel->startCol(); i < excel->endCol(); i++) {
        addColName(excel->columnIntToString(i));
    }
}

void NumberBit::modifyValue(QRegExp *rx, QExcel *excel, bool isAllCol)
{
    int startRow = excel->startRow() + 1; // 跳过标题栏
    int startCol = excel->startCol();
    int endRow = excel->endRow();
    int endCol = excel->endCol();

    for(int col = startCol; col <= endCol; col++) {
        // 如果没有全选列，且在索引集找不到，则跳过
        if(isAllCol || colIndexs().contains(col - 1)) {

            for(int row = startRow; row <= endRow; row++) {
                auto value = excel->getCellValue(row, col);

                // 跳过不是数字的列
                if(rx->exactMatch(value.toString()) == false)
                    break;

                auto intValue = value.toInt();
                auto doubleValue = value.toDouble();

                if(intValue != doubleValue) { // 不相等证明是浮点数
//                    if(excel->getFontColor(row, col) == Qt::red) {
//                        doubleValue = -qAbs(doubleValue);
//                    }

                    auto setValue = QString::number(doubleValue, 'd', 2);
                    excel->setCellString(row, col, setValue);

                    // 设置单元格变成数值，保留两位小数
//                    excel->setCellForm(row, col, excel->getNumberBitForm(2));
                } else {
                    excel->setCellString(row, col, QString::number(intValue));
                }

                if(doubleValue < 0) { // 如果原本数值小于0，给设置成红色字体
//                    excel->setFontColor(row, col, Qt::red);
                }
            }
        }
    }
}

void NumberBit::modifySheet(QRegExp *rx, QExcel *excel, bool isAllSheet, bool isAllCol)
{
    auto count = excel->getSheetsCount();
    for(int i = 0; i < count; i++) {
        // 如果没有全选sheet，且在索引集找不到，则跳过
        if(isAllSheet || workSheetIndexs().contains(i))
            excel->selectSheet(i + 1);
        // 修改值
        modifyValue(rx, excel, isAllCol);
    }
}

void NumberBit::modify(bool isAllDoc, bool isAllBook,
                       bool isAllSheet, bool isAllCol)
{
    QRegExp rx("^(-?\\d+)(\\.\\d+)?$");

    for(int i = 0; i < excels.size(); i++) {



        auto excelKey = excels.keys()[i];
        auto excel = excels[excelKey];

        // 如果全选文档，或索引集找到该文档索引
        if(isAllDoc || documentIndexs().contains(i))
            modifySheet(&rx, excel, isAllSheet, isAllCol);// 修改sheet

        // 保存文件
        QFileInfo info(excelKey);
        auto path = info.absolutePath() + "//back";
        QDir dir(path);
        if(!dir.exists())
            dir.mkdir(path);
        excel->saveAs(path + "//new_" + documentNames()[i]);
    }

}


