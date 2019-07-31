#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QAxObject>
#include <QFileDialog>
#include <QMessageBox>
#include <QTabWidget>
#include <QTableWidgetItem>
#include <QStandardItemModel>
#include <QtDebug>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_btnBrowse_clicked();

private:
    Ui::MainWindow *ui;
    QStandardItemModel* table_model;
    QString dir;
};

#endif // MAINWINDOW_H
