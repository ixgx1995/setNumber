import QtQuick 2.12
import QtQuick.Controls 2.12
import QtQuick.Window 2.12
import QtQuick.Layouts 1.12
import Qt.labs.platform 1.0
import Qt.labs.settings 1.1
import SetNumber 1.0

Window {
    id: window
    width: 640
    height: 480
    visible: true

    title: qsTr("Excel App")

    NumberBit {
        id: numberBit
        onDocumentNamesChanged: {
            var models = []
            for(var i = 0; i < documentNames.length; i++) {
                models.push({
                                check: false,
                                propName: documentNames[i]
                            });
            }
            documentList.items = models
        }

        onBookNamesChanged: {
            var models = []
            for(var i = 0; i < bookNames.length; i++) {
                models.push({
                                check: false,
                                propName: bookNames[i]
                            });
            }
            workBookList.items = models
        }

        onSheetNamesChanged: {
            var models = []
            for(var i = 0; i < sheetNames.length; i++) {
                models.push({
                                check: false,
                                propName: sheetNames[i]
                            });
            }
            workSheetList.items = models
        }

        onColNamesChanged: {
            var models = []
            for(var i = 0; i < colNames.length; i++) {
                models.push({
                                check: false,
                                propName: colNames[i]
                            });
            }
            workColList.items = models
        }
    }

    function openSelIndex() {
        documentList.selIndex(documentList.items.length - 1)
        workSheetList.selIndex(0)
        workColList.selIndex(0)
    }

    FileDialog {
        id: fileOpenDlg;
        title: qsTr("打开excel文件");
        folder: settings.docOpenFolder
        fileMode: FileDialog.OpenFile
        nameFilters: [
            "excel文件 (*.xlsx *.xls)"
        ];
        onAccepted: {
            numberBit.openDocument(file.toString().slice(8))
            window.openSelIndex()
        }
    }

    Settings {
        id:settings
        property alias docOpenFolder: fileOpenDlg.file
    }

    Rectangle {
        anchors.fill: parent
        gradient: Gradient {
            GradientStop {
                position: 0
                color: "#fdfcfb"
            }

            GradientStop {
                position: 1
                color: "#e2d1c3"
            }
        }

        DropArea {
            anchors.fill: parent
            onDropped: {
                if(drop.hasUrls) {
                    for(var i = 0; i < drop.urls.length; i++) {
                        numberBit.openDocument(drop.urls[i].slice(8))
                        console.log(drop.urls[i]);
                        console.log(drop.urls[i].slice(8)); //去掉前缀：file:///
                    }
                    window.openSelIndex()
                }
            }
        }


        CheckBox {
            id: selAllDocument
            text: "全部文件"
            font.pixelSize: 13
            anchors.leftMargin: 5
            anchors.topMargin: 5
            onCheckedChanged: documentList.setAll(checked)
        }

        /*CheckBox {
            id: selAllWorkBook
            text: "全部Book"
            font.pixelSize: 13
            y: 5 + documentList.height
            anchors.leftMargin: 5
            onCheckedChanged: workBookList.setAll(checked)
        }*/

        CheckBox {
            id: selAllSheet
            text: "全部Sheet"
            font.pixelSize: 13
            y: 5 + documentList.height + 10
            anchors.leftMargin: 5
            onCheckedChanged: workSheetList.setAll(checked)
        }

        CheckBox {
            id: selAllCol
            text: "全部列"
            font.pixelSize: 13
            y: 5 + documentList.height + workSheetList.height + 10 * 2
            anchors.leftMargin: 5
            onCheckedChanged: workColList.setAll(checked)
        }

        CheckBox {
            id: selAll
            text: "全选"
            font.pixelSize: 13
            y: 5 + documentList.height + workSheetList.height + workColList.height + 10 * 3
            anchors.leftMargin: 5
            onCheckedChanged: {
                selAllDocument.checked = selAll.checked
                selAllSheet.checked = selAll.checked
                selAllCol.checked = selAll.checked
            }
        }

        Column {
            id: proptyColumn
            anchors.fill: parent
            anchors.topMargin: 5
            anchors.leftMargin: 110
            anchors.rightMargin: 5
            anchors.bottomMargin: 100
            spacing: 10

            ProptyList {
                id: documentList
                anchors.left: parent.left
                anchors.right: parent.right
                height: parent.height / 3
                onSelItem: {
                    var index = -2;
                    var indexs = [];
                    for(var i = 0; i < items.length; i++) {
                        if(items[i].check) {
                            index = index === -2 ? i : -1;
                            indexs.push(i);
                        }
                    }

                    numberBit.documentIndexs = indexs;
                    // 全选或多选
                    if(index === -1 || index === -2) {
                        numberBit.selDocument("")
                    } else { // 单选
                        numberBit.selDocument(items[index].propName)
                    }
                }
            }

            /*ProptyList {
                id: workBookList
                anchors.left: parent.left
                anchors.right: parent.right
                height: 50
                isHor: true
                itemWidth: 100
                onSelItem: {
                    var index = -2;
                    var indexs = []
                    for(var i = 0; i < items.length; i++) {
                        if(items[i].check) {
                            index = index === -2 ? i : -1;
                            indexs.push(i);
                        }
                    }

                    numberBit.workBookIndexs = indexs;
                    // 全选或多选
                    if(index === -1 || index === -2) {
                        numberBit.selWorkBook("")
                    } else {
                        numberBit.selWorkBook(items[index].propName)
                    }

                }
            }*/

            ProptyList {
                id: workSheetList
                anchors.left: parent.left
                anchors.right: parent.right
                height: 50
                isHor: true
                itemWidth: 100
                onSelItem: {
                    var index = -2;
                    var indexs = []
                    for(var i = 0; i < items.length; i++) {
                        if(items[i].check) {
                            index = index === -2 ? i : -1;
                            indexs.push(i);
                        }
                    }

                    numberBit.workSheetIndexs = indexs;
                    // 全选或多选
                    if(index === -1 || index === -2) {
                        numberBit.selWorkSheet("")
                    } else {
                        numberBit.selWorkSheet(items[index].propName)
                    }
                }
            }

            ProptyList {
                id: workColList
                anchors.left: parent.left
                anchors.right: parent.right
                height: 50
                isHor: true
                itemWidth: 100
                onSelItem: {
                    var indexs = []
                    for(var i = 0; i <items.length; i++) {
                        if(items[i].check) {
                            indexs.push(i);
                        }
                    }
                    numberBit.colIndexs = indexs;
                }
            }
        }

        Row {
            id: row1
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.top: proptyColumn.bottom
            height: 40
            ToolButton {
                background: Rectangle {
                    id: openBack
                    signal checkColor(int check);
                    color: "#9fab54"
                    onCheckColor: color = check ? "#9fbb54" : "#9fab54"
                }
                width: parent.width
                height: parent.height
                text: "打开"
                font.pixelSize: 25
                MouseArea {
                    anchors.fill: parent
                    onPressed: openBack.checkColor(true)
                    onReleased: openBack.checkColor(false)
                    onClicked: fileOpenDlg.open()
                }
            }
        }

        Row {
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.top: row1.bottom
            anchors.bottom: parent.bottom
            ToolButton {
                background: Rectangle {
                    id: midBack
                    signal checkColor(int check);
                    color: "#82c179"
                    onCheckColor: color = check ? "#82d179" : "#82c179"
                }
                width: parent.width
                height: parent.height
                text: "处理"
                font.pixelSize: 25
                MouseArea {
                    anchors.fill: parent
                    onPressed: midBack.checkColor(true)
                    onReleased: midBack.checkColor(false)
                    onClicked: {
                        window.title = "处理中"
                        numberBit.modify(selAllDocument.checked, true,
                                         selAllSheet.checked, selAllCol.checked);
                        window.title = "处理完成"
                    }
                }
            }
        }

        ToolButton {
            text: "Test"
            visible: false
            onClicked: {
                workBookList.items.push({
                                            propName: "AAA"
                                        });
                workBookList.items.push({
                                            propName: "BBB"
                                        });
                workBookList.items = workBookList.items;
            }
        }

    }
}
