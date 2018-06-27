#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "qaxobject.h"
#include "qthread.h"

namespace Ui {
class MainWindow;
}


class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    void saveData();
    void searchData();
    void queryData();
    void Init();
    void addData();
    int getExcelRow();
    int flag=0;
    int countDown=15;
    ~MainWindow();

private:
    Ui::MainWindow *ui;
    QAxObject *workbook,*worksheet,*usedrange;

private slots:
    void on_qbt_add_clicked();//按钮函数
    void on_qbt_search_clicked();//按钮函数
    void updateCountDown();
//    void on_qlb_display_meaning_clicked();


};
#endif // MAINWINDOW_H
