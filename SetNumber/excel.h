#ifndef EXCEL_H
#define EXCEL_H

#include <QObject>
#include <QColor>
#include <QAxObject>

class QExcel : public QObject
{
    Q_OBJECT
public:
    QExcel(QString name = "");
    ~QExcel();

    void writeTitle(const QString& module, const QStringList& titles);
    void writeLineTest(int row, QVariantList pValues, quint32 valCnt);
    QVariant readLine(int row, int column);

    QAxObject* getWorkBooks();
    QAxObject* getWorkBook();
    QAxObject* getWorkSheets();
    QAxObject* getWorkSheet();

    int getWorkBookCount();
    void selectWorkBook(QString name);
    void selectWorkBook(int index);

    /**************************************************************************/
    /* 工作表                                                                 */
    /**************************************************************************/
    void selectSheet(const QString& sheetName);
    //sheetIndex 起始于 1
    void selectSheet(int sheetIndex);
    void deleteSheet(const QString& sheetName);
    void deleteSheet(int sheetIndex);
    void insertSheet(QString sheetName);
    int getSheetsCount();
    //在 selectSheet() 之后才可调用
    QString getSheetName();
    QString getSheetName(int sheetIndex);
    //datas是二维数组,datas的值类型是QList<QVariant>
    bool readSheet(QString sheetName, QVariantList& datas);
    bool writeSheet(QString sheetName, int startRow, int startCol, QVariantList& datas);
    bool writeSheet(QString sheetName, const int& startRow, const int& startCol, const int& endRow, const int& endCol, const QVariant& datas);

    /**************************************************************************/
    /* 单元格                                                                 */
    /**************************************************************************/
    void setCellString(int row, int column, const QString& value);
    //cell 例如 "A7"
    void setCellString(const QString& cell, const QString& value);
    //range 例如 "A5:C7"
    void mergeCells(const QString& range);
    void mergeCells(int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn);
    QVariant getCellValue(int row, int column);
    void clearCell(int row, int column);
    void clearCell(const QString& cell);
    void setCellDropItems(int row, int col, const QString& items);
    QString getCellForm(int row, int col);
    void setCellForm(int row, int col, QString form);
    void setColForm(int col, QString form);
    QString getGeneralForm();
    QString getTextForm();
    QString getNumberBitForm(int bit);

    /**************************************************************************/
    /* 布局格式                                                               */
    /**************************************************************************/
    void getUsedRange(int* topLeftRow, int* topLeftColumn, int* bottomRightRow, int* bottomRightColumn);
    void setColumnWidth(int column, int width);
    void setRowHeight(int row, int height);
    void setCellTextCenter(int row, int column);
    void setCellTextCenter(const QString& cell);
    void setCellTextWrap(int row, int column, bool isWrap);
    void setCellTextWrap(const QString& cell, bool isWrap);
    void setAutoFitRow(int row);
    void mergeSerialSameCellsInAColumn(int column, int topRow);
    int startRow();
    int endRow();
    int startCol();
    int endCol();
    int getUsedRowsCount();
    int getUsedColCount();
    void setCellFontBold(int row, int column, bool isBold);
    void setCellFontBold(const QString& cell, bool isBold);
    void setCellFontSize(int row, int column, int size);
    void setCellFontSize(const QString& cell, int size);
    QColor getBackColor(int row, int col);
    void setBackColor(int row, int col, QColor color);
    QColor getBoderColor(int row, int col);
    void setBoderColor(int row, int col, QColor color);
    QColor getFontColor(int row, int col);
    void setFontColor(int row, int col, QColor color);
    QColor getExcelColor(QColor qtColor);
    QString getQtColorStr(quint32 value);

    //复制指令区域的内容到目标区域，带格式
    void copyRangeToRange(QString resource, QString target);

    /**************************************************************************/
    /* 文件                                                                   */
    /**************************************************************************/
    void save();
    void saveAs(const QString& filePath);
    void close();


    /**************************************************************************/
    /* 自定义辅助方法								                                     */
    /**************************************************************************/
    QString getRangeString(int startRow, int startCol, int endRow, int endCol);

    static bool isFileUsed(const QString& fpath);
    QString columnIntToString(int col);
private:
    void freeSheet();

private:
    QAxObject* excel;
    QAxObject* workBooks;
    QAxObject* workBook;
    QAxObject* sheets;
    QAxObject* sheet;
};

#endif // EXCEL_H
