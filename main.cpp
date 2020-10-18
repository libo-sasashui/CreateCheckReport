#include "CreateCheckReport.h"
#include <QtWidgets/QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    CreateCheckReport w;
    w.show();
    return a.exec();
}
