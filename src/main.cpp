#include <QApplication>
#include <QDir>
#include <QFileInfo>
#include "main_window.h"

int main(int argc, char* argv[]) {
    QApplication app(argc, argv);
    app.setApplicationName("CRM_Automation");
    app.setApplicationVersion("3.0.0");

    MainWindow window;
    window.show();

    return app.exec();
}
