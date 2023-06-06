#include "product.h"
#include "ui_product.h"

#include <QtGui>
#include <QAxObject>
#include <QAxWidget>
#include <QApplication>
#include <QTextCodec>

#include <QChar>

#include <QFile>
#include <QTextStream>
#include <QDate>
#include <QString>
#include <QFileDialog>
#include <QDesktopServices>
#include <QDebug>

QString file1("C:\\example\\ExcelFile.xls");

Product::Product(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::Product)
{
    ui->setupUi(this);

    // кодировка для qDebug
    QTextCodec *russiancodec = QTextCodec::codecForName("Cp1251");
    QTextCodec::setCodecForLocale(russiancodec);

    // дата сегоднящнего дня
    ui->dateproduct->setDate(QDate::currentDate());

    // Начальное число сотрудника
    ui->spinNumber->setMinimum(1);
}

Product::~Product()
{
    delete ui;
}

void Product::on_pushButton_4_clicked()
{
    QStringList fileNames; // Список имен файлов для объединения

    // Цикл для генерации имен файлов и добавления их в список
    for (int i = 1; i <= 200; ++i)
    {
        QString fileName = "Product_" + QString::number(i) + ".txt";
        fileNames << fileName;
    }

    QString outputFileName = "PRODUCT.xlsx"; // Имя выходного файла для объединенного содержимого

    // создание Excel файла
    QAxObject* excel = new QAxObject("Excel.Application"); // Создаем объект Excel
    excel->dynamicCall("SetVisible(bool)", false); // Устанавливаем видимость Excel

    QAxObject* workbooks = excel->querySubObject("Workbooks"); // Получаем коллекцию книг

    QAxObject* workbook = workbooks->querySubObject("Add"); // Создаем новую книгу
    QAxObject* worksheets = workbook->querySubObject("Worksheets"); // Получаем коллекцию листов

    QAxObject* worksheet = worksheets->querySubObject("Item(int)", 1); // Получаем первый лист
    QAxObject* range = worksheet->querySubObject("Range(const QString&)", "A1"); // Получаем ячейку A1

    int column = 1; // Счетчик столбцов

    // Цикл для перебора файлов и записи их содержимого в ячейки
    foreach (const QString& fileName, fileNames)
    {
        QFile inputFile(fileName);

        if (inputFile.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            QTextStream inputStream(&inputFile);
            QString fileContent = inputStream.readAll(); // Читаем содержимое файла
            inputFile.close();

            // Извлекаем номер комнаты из имени файла
            QString prNumber = fileName.mid(8, fileName.indexOf('.') - 8);

            // Проверяем, находится ли введенный номер комнаты в диапазоне от 1 до 200
            bool validNumber = false;
            int prNum = prNumber.toInt(&validNumber);
            if (!validNumber || prNum < 1 || prNum > 200)
            {
                // Показываем сообщение об ошибке или обрабатываем недопустимый номер комнаты
                continue; // Пропускаем этот файл и переходим к следующему
            }

            // Записываем имя файла в ячейку
            range = worksheet->querySubObject("Range(const QString&)", getColumnLabel(column) + "1");
            range->setProperty("Value", fileName);

            // Обрабатываем содержимое файла и записываем его в ячейки
            QStringList lines = fileContent.split('\n');
            int row = 2; // Начинаем со следующей доступной строки

            foreach (const QString& line, lines)
            {
                QString key = line.section(':', 0, 0).trimmed(); // Извлекаем ключ перед символом ":"
                QString value = line.section(':', 1).trimmed(); // Извлекаем значение после символа ":"

                if (!key.isEmpty() && !value.isEmpty())
                {
                    // Записываем ключ и значение в соответствующие столбцы
                    range = worksheet->querySubObject("Range(const QString&)", getColumnLabel(column) + QString::number(row));
                    range->setProperty("Value", key);

                    range = worksheet->querySubObject("Range(const QString&)", getColumnLabel(column + 1) + QString::number(row));
                    range->setProperty("Value", value);

                    row++;
                    // Переходим к следующей строке для записи второй информации
                    range = worksheet->querySubObject("Range(const QString&)", getColumnLabel(column) + QString::number(row));
                }
            }

            // Увеличиваем счетчик столбцов и добавляем пустой столбец между файлами
            column += 2;
        }
    }

    workbook->dynamicCall("SaveAs(const QString&)", outputFileName); // Сохраняем книгу
    workbook->dynamicCall("Close()"); // Закрываем книгу
    excel->dynamicCall("Quit()"); // Закрываем Excel

    delete excel;

    qDebug() << "Файлы объеденены и сохранены в" << outputFileName;
}

QString Product::getColumnLabel(int column)
{
    QString label;
    while (column > 0)
    {
        int remainder = (column - 1) % 26;
        label.prepend(QChar('A' + remainder));
        column = (column - 1) / 26;
    }
    return label;
}

// Функция нажатия на кнопку "View"
void Product::on_on_View_clicked_clicked()
{
    QString fileName = QFileDialog::getOpenFileName(this, tr("Open File"), "", tr("Text Files (*.txt)"));
    if (!fileName.isEmpty())
    {
        QDesktopServices::openUrl(QUrl::fromLocalFile(fileName));

        QFile file(fileName);
        // только для чтения и в виде текстового файла
        if (file.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            QTextStream in(&file);

            // отображение в консоли
            while (!in.atEnd())
            {
                QString line = in.readLine();
                qDebug().noquote() << line;
            }
            file.close();
        }
    }
}


void Product::on_pushButton_3_clicked()
{
    ui->spinNumber->clear();
    ui->Nameproduct->clear();
    ui->priceproduct->clear();
    ui->dateproduct->setDate(QDate::currentDate());
    reject();
}




void Product::on_pushButton_clicked()
{
    QString Number = ui->spinNumber->text();
    QString Name = ui->Nameproduct->text();
    QDate Date = ui->dateproduct->date();
    QString Price = ui->priceproduct->text();

    bool validNumber = false;
    int prNumber = Number.toInt(&validNumber);
    if (!validNumber || prNumber < 1)
    {
        return;
    }

    // создание файла .txt с номером товара
    QString fileName = "Product_" + QString::number(prNumber) + ".txt";
    QFile file(fileName);
    if (file.open(QIODevice::WriteOnly | QIODevice::Append | QIODevice::Text))
    {
        // вывод в текстовый документ
        QTextStream stream(&file);
        stream << "Number: " << Number << endl;
        stream << "Name: " << Name << endl;
        stream << "Date:" << Date.toString("dd.MM.yyyy") << endl;
        stream << "Price: " << Price << endl;
        stream << endl;
        file.close();

        // очистка
        ui->spinNumber->clear();
        ui->Nameproduct->clear();
        ui->priceproduct->clear();
        ui->dateproduct->setDate(QDate::currentDate());

    }
}

