#ifndef NUMBERBIT_H
#define NUMBERBIT_H

#include <QObject>
#include <QMap>

class NumberBitPrivate;
class QExcel;

class NumberBit : public QObject
{
    Q_OBJECT
    Q_PROPERTY(QList<QString> documentNames READ documentNames WRITE setDocumentNames NOTIFY documentNamesChanged)
    Q_PROPERTY(QList<int> documentIndexs READ documentIndexs WRITE setdocumentIndexs NOTIFY documentIndexsChanged)
    Q_PROPERTY(QList<QString> bookNames READ bookNames WRITE setBookNames NOTIFY bookNamesChanged)
    Q_PROPERTY(QList<int> workBookIndexs READ workBookIndexs WRITE setWorkBookIndexs NOTIFY workBookIndexsChanged)
    Q_PROPERTY(QList<QString> sheetNames READ sheetNames WRITE setSheetNames NOTIFY sheetNamesChanged)
    Q_PROPERTY(QList<int> workSheetIndexs READ workSheetIndexs WRITE setWorkSheetIndexs NOTIFY workSheetIndexsChanged)
    Q_PROPERTY(QList<QString> colNames READ colNames WRITE setColNames NOTIFY colNamesChanged)
    Q_PROPERTY(QList<int> colIndexs READ colIndexs WRITE setColIndexs NOTIFY colIndexsChanged)
    Q_PROPERTY(int numberBit READ numberBit WRITE setNumberBit NOTIFY numberBitChanged)
    Q_DECLARE_PRIVATE(NumberBit)

public:
    NumberBit(QObject *parent = nullptr);
    ~NumberBit();

    QList<QString> documentNames() const;
    void setDocumentNames(QList<QString> value);
    void addDocumentName(QString value);
    void delDocumentName(QString value);
    void cleDocumentName();

    QList<int> documentIndexs() const;
    void setdocumentIndexs(QList<int> value);

    QList<QString> bookNames() const;
    void setBookNames(QList<QString> value);
    void addBookName(QString value);
    void delBookName(QString value);
    void cleBookName();

    QList<int> workBookIndexs() const;
    void setWorkBookIndexs(QList<int> value);

    QList<QString> sheetNames() const;
    void setSheetNames(QList<QString> value);
    void addSheetName(QString value);
    void delSheetName(QString value);
    void cleSheetName();

    QList<int> workSheetIndexs() const;
    void setWorkSheetIndexs(QList<int> value);

    QList<QString> colNames() const;
    void setColNames(QList<QString> value);
    void addColName(QString value);
    void delColName(QString value);
    void cleColName();

    QList<int> colIndexs() const;
    void setColIndexs(QList<int> value);

    int numberBit() const;
    void setNumberBit(int value);

    void closeDocument();

    void modifyValue(QRegExp *rx, QExcel *excel, bool isAllCol);
    void modifySheet(QRegExp *rx, QExcel *excel, bool isAllSheet, bool isAllCol);

    Q_INVOKABLE void openDocument(QString fpath);
    Q_INVOKABLE void selDocument(QString documentName);
    Q_INVOKABLE void selWorkBook(QString bookName);
    Q_INVOKABLE void selWorkSheet(QString sheetName);
    Q_INVOKABLE void modify(bool isAllDoc, bool isAllBook,
                            bool isAllSheet, bool isAllCol);

signals:
    void documentNamesChanged(QList<QString> value);
    void documentIndexsChanged(QList<int> value);
    void bookNamesChanged(QList<QString> value);
    void workBookIndexsChanged(QList<int> value);
    void sheetNamesChanged(QList<QString> value);
    void workSheetIndexsChanged(QList<int> value);
    void colNamesChanged(QList<QString> value);
    void colIndexsChanged(QList<int> value);
    void rowIndexChanged(int value);
    void colIndexChanged(int value);
    void numberBitChanged(int value);

protected:
    QScopedPointer<NumberBitPrivate> d_ptr;

private:
    QMap<QString, QExcel*> excels;
    QExcel* actExcel;

};

#endif  // NUMBERBIT_H
