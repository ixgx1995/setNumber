#include "excel.h"
#include <QDir>
#include <QDebug>

QExcel::QExcel(QString name) :
    excel(nullptr),
    workBooks(nullptr),
    workBook(nullptr),
    sheets(nullptr),
    sheet(nullptr)
{
    excel = new QAxObject(this);
    // 连接Excel控件，如果没有office,就用wps
    if(!excel->setControl("Excel.Application")) {
        excel->setControl("ket.Application");
    }
    excel->dynamicCall("SetUserControl(bool UserControl)", true);
    excel->dynamicCall("SetVisible (bool Visible)", "false");
    excel->setProperty("DisplayAlerts", false);
    workBooks = excel->querySubObject("Workbooks");
    if(!name.isEmpty())
        workBooks->dynamicCall("Open(const QString&)", name);
    else
        workBooks->dynamicCall("Add");
    workBook = excel->querySubObject("ActiveWorkBook");
    sheets = workBook->querySubObject("WorkSheets");
}

QExcel::~QExcel()
{
    close();
}

void QExcel::writeTitle(const QString& module, const QStringList& titles)
{
    freeSheet();

    sheets->querySubObject("Add()");
    sheet = sheets->querySubObject("Item(int)", 1);
    sheet->setProperty("Name", module);

    for (int column =0; column < titles.size(); ++column)
    {
        QAxObject* cell = sheet->querySubObject("Cells(int,int)", 1, column + 1);
        cell->dynamicCall("SetValue(const QString&)", titles.at(column));
        delete cell;
        cell = nullptr;
    }
}

void QExcel::writeLineTest(int row, QVariantList pValues, quint32 valCnt)
{
    for (quint32 column = 0; column < valCnt; ++column)
    {
        QAxObject* cell = sheet->querySubObject("Cells(int,int)", row, column + 1);
        cell->dynamicCall("SetValue(const QString&)", pValues[column]);
        delete cell;
        cell = nullptr;
    }
}

QVariant QExcel::readLine(int row, int column)
{
    QVariant data;
    QAxObject* range = sheet->querySubObject("Cells(int,int)", row, column);
    if (range)
    {
        data = range->dynamicCall("Value2()");
    }
    return data;
}

void QExcel::close()
{
    //关闭excel
    workBook->dynamicCall("Close(Boolean)", true);
    excel->dynamicCall("Quit()");

    delete sheet;
    delete sheets;
    delete workBook;
    delete workBooks;
    delete excel;

    excel = nullptr;
    workBooks = nullptr;
    workBook = nullptr;
    sheets = nullptr;
    sheet = nullptr;
}

QAxObject* QExcel::getWorkBooks()
{
    return workBooks;
}

QAxObject* QExcel::getWorkBook()
{
    return workBook;
}

QAxObject* QExcel::getWorkSheets()
{
    return sheets;
}

QAxObject* QExcel::getWorkSheet()
{
    return sheet;
}

int QExcel::getWorkBookCount()
{
    return workBooks->property("Count").toInt();
}

void QExcel::selectWorkBook(QString name)
{
    delete sheet;
    delete sheets;
    delete workBook;
    workBook = nullptr;
    sheets = nullptr;
    sheet = nullptr;

    workBook = workBooks->querySubObject("Item(const QString&)", name);
    sheets = workBook->querySubObject("WorkSheets");
}

void QExcel::selectWorkBook(int index)
{
    delete sheet;
    delete sheets;
    delete workBook;

    workBook = nullptr;
    sheets = nullptr;
    sheet = nullptr;

    workBook = workBooks->querySubObject("Item(int)", index);
    sheets = workBook->querySubObject("WorkSheets");
}

void QExcel::selectSheet(const QString& sheetName)
{
    sheet = sheets->querySubObject("Item(const QString&)", sheetName);
}

void QExcel::deleteSheet(const QString& sheetName)
{
    QAxObject* deleteSheet = sheets->querySubObject("Item(const QString&)", sheetName);
    deleteSheet->dynamicCall("delete");
}

void QExcel::deleteSheet(int sheetIndex)
{
    QAxObject* deleteSheet = sheets->querySubObject("Item(int)", sheetIndex);
    deleteSheet->dynamicCall("delete");
}

void QExcel::selectSheet(int sheetIndex)
{
    sheet = sheets->querySubObject("Item(int)", sheetIndex);
}

void QExcel::setCellString(int row, int column, const QString& value)
{
    QAxObject* range = sheet->querySubObject("Cells(int,int)", row, column);
    range->dynamicCall("SetValue(const QString&)", value);
}

void QExcel::setCellFontBold(int row, int column, bool isBold)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Bold", isBold);
}

void QExcel::setCellFontSize(int row, int column, int size)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Size", size);
}

void QExcel::copyRangeToRange(QString resource, QString target)
{
    if (sheet)
    {
        QAxObject* resRange = sheet->querySubObject("Range(const QString&)", resource);
        resRange->dynamicCall("Copy");//复制指令区域内容到剪贴板

        QAxObject* tarRange = sheet->querySubObject("Range(const QString&)", target);//选中目标区域
        tarRange->dynamicCall("PasteSpecial");
    }
}

void QExcel::mergeCells(const QString& cell)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->setProperty("MergeCells", true);
}

void QExcel::mergeCells(int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn)
{
    QString cell;
    cell.append(QChar(topLeftColumn - 1 + 'A'));
    cell.append(QString::number(topLeftRow));
    cell.append(":");
    cell.append(QChar(bottomRightColumn - 1 + 'A'));
    cell.append(QString::number(bottomRightRow));

    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->setProperty("MergeCells", true);
}

QVariant QExcel::getCellValue(int row, int column)
{
    QAxObject* range = sheet->querySubObject("Cells(int,int)", row, column);
    return range->dynamicCall("value");
}

void QExcel::save()
{
    workBook->dynamicCall("Save()");
}

void QExcel::saveAs(const QString& filePath)
{
    workBook->dynamicCall("SaveAs(const QString &)",
    QDir::toNativeSeparators(filePath));
}

int QExcel::getSheetsCount()
{
    return sheets->property("Count").toInt();
}

QString QExcel::getSheetName()
{
    return sheet->property("Name").toString();
}

QString QExcel::getSheetName(int sheetIndex)
{
    QAxObject* a = sheets->querySubObject("Item(int)", sheetIndex);
    return a->property("Name").toString();
}

bool QExcel::readSheet(QString sheetName, QVariantList& datas)
{
    selectSheet(sheetName);
    if (sheet == nullptr)
        return false;
    if (QAxObject* usedRange = sheet->querySubObject("UsedRange"))
    {
        QVariantList varRows = usedRange->dynamicCall("Value").toList();
        datas = varRows;
    }
    return true;
}

bool QExcel::writeSheet(QString sheetName, int startRow, int startCol, QVariantList& datas)
{
    selectSheet(sheetName);
    if (sheet == nullptr)
        return false;

    for (auto iter : datas)
    {
        QList<QVariant> rowDatas = iter.toList();
        int curCol = startCol;
        for (auto value : rowDatas)
        {
            setCellString(startRow, curCol, value.toString());
            ++curCol;
        }
        ++startRow;
    }

    return true;
}

bool QExcel::writeSheet(QString sheetName, const int& startRow, const int& startCol, const int& endRow, const int& endCol,
    const QVariant& datas)
{
    selectSheet(sheetName);
    if (sheet == nullptr)
        return false;
    QString rangeStr = getRangeString(startRow, startCol, endRow, endCol);
    QAxObject* range = sheet->querySubObject("Range(const QString&)", rangeStr);
    if (range)
    {
        bool res = range->setProperty("Value", datas);
        return res;
    }
    return true;
}

void QExcel::getUsedRange(int* topLeftRow, int* topLeftColumn, int* bottomRightRow, int* bottomRightColumn)
{
    QAxObject* usedRange = sheet->querySubObject("UsedRange");
    *topLeftRow = usedRange->property("Row").toInt();
    *topLeftColumn = usedRange->property("Column").toInt();

    QAxObject* rows = usedRange->querySubObject("Rows");
    *bottomRightRow = *topLeftRow + rows->property("Count").toInt() - 1;

    QAxObject* columns = usedRange->querySubObject("Columns");
    *bottomRightColumn = *topLeftColumn + columns->property("Count").toInt() - 1;
}

void QExcel::setColumnWidth(int column, int width)
{
    QString columnName;
    columnName.append(QChar(column - 1 + 'A'));
    columnName.append(":");
    columnName.append(QChar(column - 1 + 'A'));

    QAxObject* col = sheet->querySubObject("Columns(const QString&)", columnName);
    col->setProperty("ColumnWidth", width);
}

void QExcel::setCellTextCenter(int row, int column)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("HorizontalAlignment", -4108);//xlCenter
}

void QExcel::setCellTextWrap(int row, int column, bool isWrap)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("WrapText", isWrap);
}

void QExcel::setAutoFitRow(int row)
{
    QString rowsName;
    rowsName.append(QString::number(row));
    rowsName.append(":");
    rowsName.append(QString::number(row));

    QAxObject* rows = sheet->querySubObject("Rows(const QString &)", rowsName);
    rows->dynamicCall("AutoFit()");
}

void QExcel::insertSheet(QString sheetName)
{
    sheets->querySubObject("Add()");
    QAxObject* a = sheets->querySubObject("Item(int)", 1);
    a->setProperty("Name", sheetName);
}

void QExcel::mergeSerialSameCellsInAColumn(int column, int topRow)
{
    int a, b, c, rowsCount;
    getUsedRange(&a, &b, &rowsCount, &c);

    int aMergeStart = topRow, aMergeEnd = topRow + 1;

    QString value;
    while (aMergeEnd <= rowsCount)
    {
        value = getCellValue(aMergeStart, column).toString();
        while (value == getCellValue(aMergeEnd, column).toString())
        {
            clearCell(aMergeEnd, column);
            aMergeEnd++;
        }
        aMergeEnd--;
        mergeCells(aMergeStart, column, aMergeEnd, column);

        aMergeStart = aMergeEnd + 1;
        aMergeEnd = aMergeStart + 1;
    }
}

void QExcel::clearCell(int row, int column)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    clearCell(cell);
}

void QExcel::clearCell(const QString& cell)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("ClearContents()");
}

void QExcel::setCellDropItems(int row, int col, const QString& items)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", getRangeString(row, col, row, col));
    QAxObject* validTion = range->querySubObject("Validation");
    validTion->dynamicCall("Modify(int,int,int,QVariant)", 3, 1, 1, items);
    validTion->setProperty("IgnoreBlank", true);
    validTion->setProperty("InCellDropdown", true);
    validTion->setProperty("InputTitle", "test");
}

QString QExcel::getCellForm(int row, int col)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", getRangeString(row, col, row, col));
    auto str = range->property("NumberFormat").toString();
    return str;
}

void QExcel::setCellForm(int row, int col, QString form)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", getRangeString(row, col, row, col));
    range->setProperty("NumberFormat", form);
}

void QExcel::setColForm(int col, QString form)
{
    QAxObject* column = sheet->querySubObject("Columns(int)", col);
    column->setProperty("NumberFormat", form);
}

QString QExcel::getGeneralForm()
{
    return "General";
}

QString QExcel::getTextForm()
{
    return "@";
}

QString QExcel::getNumberBitForm(int bit)
{
    QString str("0.");
    while(bit--) {
        str.append("0");
    }
    return str.append("_ ");
}

int QExcel::startRow()
{
    QAxObject* usedrange = sheet->querySubObject("UsedRange");
    return usedrange->property("Row").toInt();
}

int QExcel::endRow()
{
    QAxObject* usedRange = sheet->querySubObject("UsedRange");
    int startRow = usedRange->property("Row").toInt();
    QAxObject* rows = usedRange->querySubObject("Rows");
    return startRow + rows->property("Count").toInt();
}

int QExcel::startCol()
{
    QAxObject* usedRange = sheet->querySubObject("UsedRange");
    return usedRange->property("Column").toInt();
}

int QExcel::endCol()
{
    QAxObject* usedRange = sheet->querySubObject("UsedRange");
    int startCol = usedRange->property("Column").toInt();
    QAxObject* cols = usedRange->querySubObject("Columns");
    return startCol + cols->property("Count").toInt();
}

int QExcel::getUsedRowsCount()
{
    return endRow() - 1;
}

int QExcel::getUsedColCount()
{
    return endCol() - 1;
}

void QExcel::setCellString(const QString& cell, const QString& value)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("SetValue(const QString&)", value);
}

void QExcel::setCellFontSize(const QString& cell, int size)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Size", size);
}

QColor QExcel::getBackColor(int row, int col)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", getRangeString(row, col, row, col));
    QAxObject* cells = range->querySubObject("Columns");
    QAxObject* interior = cells->querySubObject("Interior");
    auto color = getQtColorStr(interior->property("Color").toUInt());
    return getExcelColor(color);
}

void QExcel::setBackColor(int row, int col, QColor color)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", getRangeString(row, col, row, col));
    QAxObject* cells = range->querySubObject("Columns");
    QAxObject* interior = cells->querySubObject("Interior");
    interior->setProperty("Color", color);
}

QColor QExcel::getBoderColor(int row, int col)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", getRangeString(row, col, row, col));
    QAxObject* cells = range->querySubObject("Columns");
    QAxObject* borders = cells->querySubObject("Borders");
    auto color = getQtColorStr(borders->property("Color").toUInt());
    return getExcelColor(color);
}

void QExcel::setBoderColor(int row, int col, QColor color)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", getRangeString(row, col, row, col));
    QAxObject* cells = range->querySubObject("Columns");
    QAxObject* borders = cells->querySubObject("Borders");
    borders->setProperty("Color", getExcelColor(color));
}

QColor QExcel::getFontColor(int row, int col)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", getRangeString(row, col, row, col));
    QAxObject* font = range->querySubObject("Font");
    auto color = getQtColorStr(font->property("Color").toUInt());
    return getExcelColor(color);
}

void QExcel::setFontColor(int row, int col, QColor color)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", getRangeString(row, col, row, col));
    QAxObject* font = range->querySubObject("Font");
    font->setProperty("Color", color);
}

QColor QExcel::getExcelColor(QColor color)
{
    int r = 0, g = 0, b = 0;
    color.getRgb(&r, &g, &b);
    QString str("#");
    str.append(b < 16 ? "0" : "");
    str.append(QString::number(b, 16));
    str.append(g < 16 ? "0" : "");
    str.append(QString::number(g, 16));
    str.append(r < 16 ? "0" : "");
    str.append(QString::number(r, 16));
    return QColor(str);
}

QString QExcel::getQtColorStr(quint32 value)
{
    QString str = QString::number(value, 16);
    while(str.length() < 6)
        str.insert(0, "0");
    return "#" + str;
}

void QExcel::setCellTextCenter(const QString& cell)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("HorizontalAlignment", -4108);//xlCenter
}

void QExcel::setCellFontBold(const QString& cell, bool isBold)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Bold", isBold);
}

void QExcel::setCellTextWrap(const QString& cell, bool isWrap)
{
    QAxObject* range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("WrapText", isWrap);
}

void QExcel::setRowHeight(int row, int height)
{
    QString rowsName;
    rowsName.append(QString::number(row));
    rowsName.append(":");
    rowsName.append(QString::number(row));

    QAxObject* r = sheet->querySubObject("Rows(const QString &)", rowsName);
    r->setProperty("RowHeight", height);
}

QString QExcel::getRangeString(int startRow, int startCol, int endRow, int endCol)
{
    QString res = QString("%1%2:%3%4").arg(columnIntToString(startCol)).arg(startRow).arg(columnIntToString(endCol)).arg(endRow);
    if (startRow == endRow && startCol == endCol)
    {
        res = QString("%1%2").arg(columnIntToString(startCol)).arg(startRow);
    }
    return	res;
}

bool QExcel::isFileUsed(const QString& fpath)
{
    bool isUsed = false;

    QString fpathx = fpath + "x";

    QFile file(fpath);
    if (file.exists())
    {
        bool isCanRename = file.rename(fpath, fpathx);
        if (isCanRename == false)
        {
            isUsed = true;
        }
        else
        {
            file.rename(fpathx, fpath);
        }
    }
    file.close();

    return isUsed;
}

void QExcel::freeSheet()
{
    if (sheet)
    {
        delete sheet;
        sheet = nullptr;
    }
}

QString QExcel::columnIntToString(int col)
{
    QString res;
    if (col / 26 > 0)
    {
        res += columnIntToString(col - 26);
    }
    else
    {
        res = QString("%1").arg(QString(col - 1 + 'A'));
    }
    return res;
}
