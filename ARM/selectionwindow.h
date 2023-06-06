#ifndef SELECTIONWINDOW_H
#define SELECTIONWINDOW_H

#include "info.h"
#include "product.h"

#include <QDialog>

namespace Ui {
class authorization;
}

class authorization : public QDialog
{
    Q_OBJECT

public:
    explicit authorization(QWidget *parent = nullptr);
    ~authorization();

private:
    Ui::authorization *ui;
};

#endif // SELECTIONWINDOW_H
