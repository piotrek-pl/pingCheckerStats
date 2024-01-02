#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QFile>
#include <QMainWindow>
#include <QSpinBox>
#include "xlsxdocument.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_openFileButton_clicked();

    void on_generateExcelFileButton_clicked();

    //void updateSeconds();

private:
    Ui::MainWindow *ui;
    QString fileName;
    QFile file;

    int interval;
    void processAndWriteLine(const QString &line, const QDateTime &time, int row, QXlsx::Document &xlsx, const QXlsx::Format &fontFormat);
};
#endif // MAINWINDOW_H
