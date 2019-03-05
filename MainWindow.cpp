#include "MainWindow.h"
#include "ui_MainWindow.h"
#include <QtMsgHandler>
#include <QTimer>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    this->CloseOffice();
    delete ui;
}

void MainWindow::OpenExcel(QString &filename)
{
    this->CloseOffice();
    officeContent_ = new QAxWidget("Excel.Application", this->ui->widget);
    officeContent_->dynamicCall("SetVisible (bool Visible)","false");//不显示窗体
    officeContent_->setProperty("DisplayAlerts", false);
    auto rect = this->ui->widget->geometry();
    officeContent_-> setGeometry(rect);
    officeContent_->setControl(filename);
    officeContent_->show();
}

void MainWindow::OpenWord(QString &filename)
{
    this->CloseOffice();
    officeContent_ = new QAxWidget("Word.Application", this->ui->widget);
    officeContent_->dynamicCall("SetVisible (bool Visible)","false");//不显示窗体
    officeContent_->setProperty("DisplayAlerts", false);
    auto rect = this->ui->widget->geometry();
    officeContent_-> setGeometry(rect);
    officeContent_->setControl(filename);
    officeContent_->show();
    //this->ui->gridLayout->addWidget(officeContent_);
}

void MainWindow::OpenPdf(QString &filename)
{
    this->CloseOffice();
    officeContent_ = new QAxWidget(this);
    if(!officeContent_->setControl("Adobe PDF Reader"))
        QMessageBox::critical(this, "Error", "没有安装pdf！");
    this->ui->gridLayout->addWidget(officeContent_);
    this->ui->widget->hide();
    officeContent_->dynamicCall(
                "LoadFile(const QString&)",
                filename);
}

void MainWindow::CloseOffice()
{
    if(this->officeContent_)
    {
        officeContent_->close();
        officeContent_->clear();
        delete officeContent_;
        officeContent_ = nullptr;
    }
}

void MainWindow::on_pushButton_clicked()
{
    QFileDialog dialog;
    dialog.setFileMode(QFileDialog::ExistingFile);
    dialog.setViewMode(QFileDialog::Detail);
    dialog.setOption(QFileDialog::ReadOnly, true);
    dialog.setWindowTitle(QString("QAXwidget操作文件"));
    dialog.setDirectory(QString("./"));
    dialog.setNameFilter(QString("所有文件(*.*);;excel(*.xlsx);;word(*.docx *.doc);;pdf(*.pdf)"));
    if (dialog.exec())
    {
        QStringList files = dialog.selectedFiles();
        for(auto filename : files)
        {
            if(filename.endsWith(".xlsx"))
            {
                this->OpenExcel(filename);
            }
            else if(filename.endsWith(".docx") || filename.endsWith(".doc"))
            {
                this->OpenWord(filename);
            }
            else if(filename.endsWith(".pdf"))
            {
                this->OpenPdf(filename);
            }
        }
    }
}
