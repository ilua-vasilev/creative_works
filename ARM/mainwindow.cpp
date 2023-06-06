#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QtGui>
#include <QAxObject>
#include <QAxWidget>
#include <QApplication>
#include <QTextCodec>
#include <QMessageBox>
#include <QChar>
#include <QPixmap>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    QPixmap pix(":/im/img/home.png");
    int w = ui->image->width();
    int h = ui->image->height();

    ui->image->setPixmap(pix.scaled(w,h,Qt::KeepAspectRatio));
}

MainWindow::~MainWindow()
{
    delete ui;
}

// проверка на совпадения логина и пароля и переход к следующему окну
void MainWindow::on_pushButton_clicked()
{
   QString Login = ui->Login->text();
   QString Password = ui->Pass->text();
   if (Login=="ARM" && Password=="123")
   {
       hide();
       selection_window *infoDialog = new selection_window(this); // Создание объекта класса Info
       infoDialog->show(); // Отображение диалогового окна
   }
   else
   {
       QMessageBox::warning(this, "Authorization", "Failed authorization attempt!!!");
   }
}

