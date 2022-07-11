#include <QGuiApplication>
#include <QQmlApplicationEngine>
#include "excel.h"
#include "numberbit.h"
#include <QDebug>

int main(int argc, char *argv[])
{
#if QT_VERSION < QT_VERSION_CHECK(6, 0, 0)
    QCoreApplication::setAttribute(Qt::AA_EnableHighDpiScaling);
#endif

    QGuiApplication app(argc, argv);

    /*
    QExcel excel("C:\\Users\\93011\\Desktop\\b.xls");
    auto count = excel.getSheetsCount();
    excel.selectSheet(1);
    auto startRow = excel.startRow();
    auto startCol = excel.startCol();
//    excel.setBackColor(1, 1, Qt::red);
//    excel.setFontColor(2, 1, Qt::red);
    auto color = excel.getFontColor(startRow + 2, startCol);
    qDebug()<<color.red();
    qDebug()<<color.green();
    qDebug()<<color.blue();
    excel.setColForm(1, excel.getNumberBitForm(2));
    excel.saveAs("C:\\Users\\93011\\Desktop\\test2.xlsx");
    excel.close();*/

    qmlRegisterType<NumberBit>("SetNumber", 1, 0, "NumberBit");

    QQmlApplicationEngine engine;
    const QUrl url(QStringLiteral("qrc:/main.qml"));
    QObject::connect(&engine, &QQmlApplicationEngine::objectCreated,
                     &app, [url](QObject *obj, const QUrl &objUrl) {
        if (!obj && url == objUrl)
            QCoreApplication::exit(-1);
    }, Qt::QueuedConnection);
    engine.load(url);

    return app.exec();
}
