#include "MainWindow.h"
#include <QApplication>
#include "Global.h"
#include <QByteArray>
int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    MainWindow w;
    w.show();
    return a.exec();
}
