#ifndef SELECTION_WINDOW_H
#define SELECTION_WINDOW_H

#include <QDialog>

#include "info.h"
#include "product.h"

namespace Ui {
class selection_window;
}

class selection_window : public QDialog
{
    Q_OBJECT

public:
    explicit selection_window(QWidget *parent = nullptr);
    ~selection_window();

private slots:
    void on_pushButton_clicked();

    void on_pushButton_3_clicked();

private:
    Ui::selection_window *ui;
    Info info;
    Product product;
};

#endif // SELECTION_WINDOW_H
