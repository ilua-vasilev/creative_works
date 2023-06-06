#include "info.h"
#include "ui_info.h"
#include <QFile>
#include <QTextStream>
#include <QDate>
#include <QString>
#include <QFileDialog>
#include <QDesktopServices>
#include <QDebug>

#include <QtGui>
#include <QAxObject>
#include <QAxWidget>
#include <QApplication>
#include <QTextCodec>

#include <QChar>

QString file2("C:\\example\\ExcelFile.xls");

Info::Info(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::Info)
{
    ui->setupUi(this);

    // дата сегоднящнего дня
    ui->Birth_date_push->setDate(QDate::currentDate());

    // Начальное число сотрудника
    ui->Number_Spin->setMinimum(1);

    //Кодировка для QDebug
    QTextCodec *russiancodec = QTextCodec::codecForName("Cp1251");
    QTextCodec::setCodecForLocale(russiancodec);

}

// деструктор
Info::~Info()
{
    delete ui;
}

// Функция нажатия на кнопку "Save"
void Info::on_Save_clicked()
{
    QString Number = ui->Number_Spin->text();
    QString Surname = ui->surname_push->text();
    QString name = ui->Name_push->text();
    QString Middlename = ui->middle_name_push->text();
    QString salary = ui->Salary_line->text();
    QDate birthDate = ui->Birth_date_push->date();

    bool validNumber = false;
    int SotNumber = Number.toInt(&validNumber);
    if (!validNumber || SotNumber < 1)
    {
        return;
    }

    // создание файла .txt с номером сотрудника
    QString fileName = "Sotrudnik_" + QString::number(SotNumber) + ".txt";
    QFile file(fileName);
    if (file.open(QIODevice::WriteOnly | QIODevice::Append | QIODevice::Text))
    {
        // вывод в текстовый документ
        QTextStream stream(&file);
        stream << "Surname: " << Surname << endl;
        stream << "Name: " << name << endl;
        stream << "Middle name: " << Middlename << endl;
        stream << "Date of birth: " << birthDate.toString("dd.MM.yyyy") << endl;
        stream << "Salary: " << salary << endl;
        stream << endl;
        file.close();

        // очистка
        ui->Number_Spin->clear();
        ui->surname_push->clear();
        ui->Name_push->clear();
        ui->middle_name_push->clear();
        ui->Salary_line->clear();
        ui->Birth_date_push->setDate(QDate::currentDate());

    }
}

// Функция нажатия на кнопку "View"
void Info::on_View_clicked()
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

// // Функция нажатия на кнопку "ОК"
void Info::on_pushButton_clicked()
{
    ui->Number_Spin->clear();
    ui->surname_push->clear();
    ui->Name_push->clear();
    ui->middle_name_push->clear();
    ui->Birth_date_push->setDate(QDate::currentDate());
    ui->Salary_line->clear();
    reject();
}

// Функция нажатия на кнопку "Save in CSV"
void Info::on_pushButton_2_clicked()
{
    QStringList fileNames; // Список имен файлов для объединения

    // Цикл для генерации имен файлов и добавления их в список
    for (int i = 1; i <= 200; ++i)
    {
        QString fileName = "Sotrudnik_" + QString::number(i) + ".txt";
        fileNames << fileName;
    }

    QString outputFileName = "SOTRUDNIKI.xlsx"; // Имя выходного файла для объединенного содержимого

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
            QString sotNumber = fileName.mid(10, fileName.indexOf('.') - 10);

            // Проверяем, находится ли введенный номер комнаты в диапазоне от 1 до 200
            bool validNumber = false;
            int sotNum = sotNumber.toInt(&validNumber);
            if (!validNumber || sotNum < 1 || sotNum > 200)
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

QString Info::getColumnLabel(int column)
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
