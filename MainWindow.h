#ifndef MAINWINDOW_H
#define MAINWINDOW_H
#pragma execution_character_set("utf-8")
#include <QMainWindow>
#include <QPlainTextEdit>
#include <QGroupBox>
#include <QVBoxLayout>
#include <QFileDialog>
#include <QAxWidget>
#include <QAxSelect>
#include <QAxObject>
#include <QTabWidget>
#include <QScrollArea>
#include <QMessageBox>

//QAxwidget打开office和pdf

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();
    void OpenExcel(QString &filename);
    void OpenWord(QString &filename);
    void OpenPdf(QString &filename);
    void CloseOffice();

private slots:

    void on_pushButton_clicked();

private:
    Ui::MainWindow *ui;
    QAxWidget* officeContent_ = nullptr;
    QAxObject* m_document = nullptr;
};

#endif // MAINWINDOW_H
