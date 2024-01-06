#include <QObject>
#include <QJsonArray>
#include <QJsonObject>
#include <QJsonDocument>
#include <QFile>
#include <QDebug>
#include <QMap>
#include "xlsxdocument.h"

#define PATH_LOCAL "/home/nct/TungNC2/Convert_Excel2Json/Data"

using namespace QXlsx;

void convertData(QString path)
{
    Document data(path);
    if (!data.load()) {
        qDebug() << "Cannot open file: " << path;
        return;
    }

    QJsonArray arr;

    int row = 4;
    do {
        QJsonObject obj;
        obj["index"] = row - 4;

        QJsonArray arrEng, arrVn, arrNote;
        for (int col = 1; col < 13; col++) {
            QString cell = data.read(row, col).toString();
            if(col < 9 && !cell.isEmpty())
                arrEng.append(cell);
            if(col >= 9 && !cell.isEmpty()) {
                arrVn.append(cell);
            }

        }
        obj["words"] = arrEng;
        obj["means"] = arrVn;
        obj["notes"] = "";
        arr.append(obj);
        row++;
    } while (data.read( row, 1).toString() != "");

    QJsonDocument jDoc(arr);
    QString jString = jDoc.toJson();

    path = PATH_LOCAL + QString("/data.json");
    QFile file(path);
    if(!file.open(QIODevice::WriteOnly | QIODevice::Text))
    {
        qDebug() << "Cannot open file: " << path;
        qDebug() << "Error: " << file.errorString();
        return;
    }
    QTextStream stream(&file);
    stream << jString;
    file.close();
    qDebug() << "Convert Excel to Json successfully";
}

int main(int argc, char *argv[])
{
    convertData(PATH_LOCAL + QString("/data.xlsx"));
    return 0;
}
