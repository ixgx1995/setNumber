import QtQuick 2.12
import QtQuick.Controls 2.12

Rectangle {
    id: propList
    color: "#8fa155"
    border.width: 1
    border.color: "#000000"

    property var items: []

    property bool isHor: false
    property var fontColor: "#00ABAB"
    property var backColor1: "#e7f0fd"
    property var backColor2: "#accbee"
    property int itemWidth: listView.width

    signal selItem(int index)
    signal setAll(bool value)

    function selIndex(index) {
        listView.checkOne(index)
    }

    Component {
        id: listDelegate

        Rectangle {
            id: item_delegate
            width: itemWidth
            height: 30
            //记录Item选中状态
            property bool checked: false
            property bool isMul: false
            border.color: backColor1
            color: item_delegate.checked ? backColor2 : backColor1
            onCheckedChanged: {
                items[index].check = checked
                selItem(index)
            }

            Label {
                color: fontColor
                anchors.rightMargin: 5
                anchors.right: parent.right
                anchors.verticalCenter: parent.verticalCenter
                text: modelData.propName
            }

            Connections {
                target: propList
                onSetAll: {
                    item_delegate.checked = value
                    items[index].check = value
                }
            }

            Connections {
                target: listView
                onCheckOne: {
                        item_delegate.checked = (idx === index);
                }
                onCheckMul: {
                    //连续多选时，判断在起始点前还是后，然后把中间的选中
                    if(idx > listView.mulBegin){
                        item_delegate.checked = (index >= listView.mulBegin && index <= idx);
                    }else{
                        item_delegate.checked = (index <= listView.mulBegin && index >= idx);
                    }
                }
            }

            MouseArea {
                id: item_mousearea
                anchors.fill: parent



                onClicked: {
                    //ctrl+多选，shift+连选，默认单选
                    switch(mouse.modifiers) {
                    case Qt.ControlModifier:
                        item_delegate.checked =! item_delegate.checked;
                        isMul = true;
                        break;
                    case Qt.ShiftModifier:
                        listView.checkMul(index);
                        isMul = true;
                        break;
                    default:
                        listView.checkOne(index);
                        listView.mulBegin = index;
                        break;
                    }
                }
            }

        }
    }

    ListView {
        id: listView

        anchors.fill: parent

        //按住shift时，连续多选的起点
        property int mulBegin: 0
        //单选信号
        signal checkOne(int idx)
        //多选信号
        signal checkMul(int idx)

        onCheckOne: listView.mulBegin = idx

        clip: true
        //取消滑动滚动
        //boundsBehavior: Flickable.StopAtBounds
        anchors.margins: 1
        spacing: 1
        model: items
        orientation: propList.isHor ? ListView.Horizontal : ListView.Vertical

        delegate: listDelegate

        ScrollBar.vertical: ScrollBar {
            visible: !isHor
        }

        ScrollBar.horizontal: ScrollBar {
            visible: isHor
        }
    }
}

