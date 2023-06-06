#include "selection_window.h"
#include "ui_selection_window.h"

selection_window::selection_window(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::selection_window)
{
    ui->setupUi(this);
}

selection_window::~selection_window()
{
    delete ui;
}

// Открытие окна заполнения формы сотрудника
void selection_window::on_pushButton_clicked()
{
    Info *infoDialog = new Info(this); // Создание объекта класса Info
    infoDialog->show(); // Отображение диалогового окна
}

//открытие окна заполнения формы товара
void selection_window::on_pushButton_3_clicked()
{
    Product *infoDialog = new Product(this); // Создание объекта класса Info
    infoDialog->show(); // Отображение диалогового окна
}

