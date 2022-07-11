import QtQuick 2.15
import QtQuick.Controls 2.15
import QtQuick.Layouts 1.12

Rectangle {
    color: "#8fa155"
    border.width: 1
    border.color: "#000000"

    property var items: []

    property var backColor1: "#e7f0fd"
    property var backColor2: "#accbee"

    Component {
        id: delegate

        RowLayout {
            anchors.left: parent.left
            anchors.right: parent.right
            spacing: 0
            Action {
                id: actionIcon1
                enabled: true
                checkable: true
                onCheckedChanged: {
                    items[index].check = checked ? true : false
                    back.checkColor(checked)
                }
            }

            ToolButton {
                id: toolButton
                Layout.fillWidth: true
                display: Button.TextUnderIcon
                action: actionIcon1

                Layout.preferredHeight: 30
                background: Rectangle {
                    id: back
                    signal checkColor(int check);
                    color: backColor1
                    onCheckColor: color = check ? backColor2 : backColor1
                }
                Label {
                    color: "#ABAB00"
                    anchors.rightMargin: 5
                    anchors.right: parent.right
                    anchors.verticalCenter: parent.verticalCenter
                    text: modelData.documentName
                }
            }
        }
    }

    ListView {
        id: list
        clip: true
        spacing: 1
        anchors.margins: 1
        anchors.fill: parent
        model: parent.items
        delegate: delegate

        ScrollBar.vertical: ScrollBar {

        }
    }
}
