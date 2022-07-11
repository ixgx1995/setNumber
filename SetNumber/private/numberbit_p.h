#ifndef NUMBERBIT_P_H
#define NUMBERBIT_P_H

#include <QObject>

class NumberBitPrivate : public QObject
{

public:
    NumberBitPrivate()

    {}

    QList<QString> documentNames;
    QList<int> documentIndexs;
    QList<QString> bookNames;
    QList<int> workBookIndexs;
    QList<QString> sheetNames;
    QList<int> workSheetIndexs;
    QList<QString> colNames;
    QList<int> colIndexs;
    int numberBit;
};

#endif  // NUMBERBIT_P_H
