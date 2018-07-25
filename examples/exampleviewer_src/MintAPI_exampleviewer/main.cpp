#include <QApplication>
#include "html5applicationviewer.h"
#include <QtConfig>
#include <QGraphicsWebView>
#include <QGraphicsItem>

QString temp_files();

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);

    Html5ApplicationViewer viewer;
    viewer.setOrientation(Html5ApplicationViewer::ScreenOrientationAuto);
    viewer.setMinimumSize(500, 480);

    viewer.showExpanded();

    viewer.webView()->setFlag(QGraphicsItem::ItemIsSelectable, false);

    viewer.loadFile(temp_files());

    return app.exec();
}

QString temp_files(){
    return QLatin1String("html/index.html");
}
