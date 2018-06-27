#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "qdebug.h"
#include "qaxobject.h"
#include "qstring.h"
#include "qtimer.h"
#include "ctime"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

//    p=new classtest();//初始化指针
//    connect(p,SIGNAL(Signal1()),this,SLOT(SlotFunction()));//绑定自定义信号与槽函数

    Init();
}

MainWindow::~MainWindow()
{
    // step6: 保存文件
    workbook->dynamicCall("Save()");  //保存文件
    workbook->dynamicCall("Close(Boolean)", false);  //关闭文件
    delete ui;
}

void MainWindow::on_qbt_add_clicked()
{

    addData();
}


void MainWindow::on_qbt_search_clicked()
{
    searchData();
}

void MainWindow::addData()
{
    queryData();
    switch (flag) {
    case 0:
        qDebug()<<tr("do not find one!");
        saveData();
        break;
    case 1:
        qDebug()<<tr("find one!");
        flag=0;
        break;
    default:
        break;
    }
}

void MainWindow::queryData()
{
    int intRow =getExcelRow();
    QAxObject* cell;
    for(int i=1;i<=intRow;i++)
    {
        cell= worksheet->querySubObject("Cells(int, int)", i, 1);  //获单元格值
        QString str=cell->dynamicCall("Value2()").toString();
        if(str.compare(ui->qle_word->text())==0)
        {
            flag=1;
            break;
        }
    }
//    flag=0;
}

void MainWindow::searchData()
{
    int intRow =getExcelRow();
    QAxObject* cell;
    bool temp=false;
    for(int i=1;i<=intRow;i++)
    {
        cell= worksheet->querySubObject("Cells(int, int)", i, 1);  //获单元格值
        QString str=cell->dynamicCall("Value2()").toString();
        if(str.compare(ui->qle_search->text())==0)
        {
            cell= worksheet->querySubObject("Cells(int, int)", i, 2);  //获单元格值
            ui->qlb_translation->clear();
            ui->qlb_translation->setText(cell->dynamicCall("Value2()").toString());
            temp=true;
            break;
        }
    }
    if(!temp)
    {
        ui->qlb_translation->clear();
        ui->qlb_translation->setText("Do not exist !");
    }
}

void MainWindow::saveData()
{

    int intRow =getExcelRow();
    intRow++;
    QString str=QString::number(intRow);
    qDebug()<<str;
    // step5: 读和写
    QAxObject* cell;
    cell = worksheet->querySubObject("Cells(int, int)", intRow, 1);  //获单元格值
    cell->dynamicCall("SetValue(conts QVariant&)", ui->qle_word->text()); // 设置单元格的值

    cell= worksheet->querySubObject("Cells(int, int)", intRow, 2);  //获单元格值
    cell->dynamicCall("SetValue(conts QVariant&)", ui->qle_meaning->text()); // 设置单元格的值

    // step6: 保存文件
    workbook->dynamicCall("Save()");  //保存文件
}

void MainWindow::Init()
{
    ui->qbt_add->setFocus(); //设置默认焦点
    ui->qbt_add->setDefault(true); //设置默认按钮，设置了这个属性，当用户按下回车的时候，就会按下该按钮

    ui->qbt_search->setFocus(); //设置默认焦点
    ui->qbt_search->setDefault(true); //设置默认按钮，设置了这个属性，当用户按下回车的时候，就会按下该按钮

    // step1：连接控件
    QAxObject* excel = new QAxObject(this);
    excel->setControl("Excel.Application");  // 连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", "false"); // 不显示窗体
    // 不显示任何警告信息。如果为true, 那么关闭时会出现类似"文件已修改，是否保存"的提示
    excel->setProperty("DisplayAlerts", false);
    // step2: 打开工作簿
    QAxObject* workbooks = excel->querySubObject("WorkBooks"); // 获取工作簿集合
    //E:\\data.xlsx
    workbook = workbooks->querySubObject("Open(const QString&)",
                                                    "E:\\WorkSpace_Qt_5.9.6\\Excel\\data.xlsx"); // 从控件lineEdit获取文件名
    // step3: 打开sheet
    worksheet = workbook->querySubObject("WorkSheets(int)", 1); // 获取工作表集合的工作表1， 即sheet1
    QAxObject *cell;
    cell = worksheet->querySubObject("Cells(int, int)", 1, 1);  //获单元格值
    cell->dynamicCall("SetValue(conts QVariant&)","words"); // 设置单元格的值

    cell= worksheet->querySubObject("Cells(int, int)", 1, 2);  //获单元格值
    cell->dynamicCall("SetValue(conts QVariant&)","meaning"); // 设置单元格的值

    QTimer *timer=new QTimer();
    connect(timer,SIGNAL(timeout()),this,SLOT(updateCountDown()));
    timer->start(1000);
}


void MainWindow::updateCountDown(){
    ui->qlb_countDown->clear();
    QString str=QString::number(countDown);
    ui->qlb_countDown->setText(str+"s");
    if(countDown==15){
        int Row=getExcelRow();
        if(Row!=1){
            qsrand(time(NULL));
            int randomRow;
            while(true){
                randomRow=qrand()%Row;
                if(randomRow!=0)
                    break;
            }
            randomRow++;
            QString str_=QString::number(randomRow);
            qDebug()<<str_+"**************";
            if(randomRow!=1){
                QAxObject* cell;  //获单元格值
                cell= worksheet->querySubObject("Cells(int, int)", randomRow, 1);  //获单元格值
                ui->qlb_display_word->clear();
                ui->qlb_display_word->setText(cell->dynamicCall("Value2()").toString());

                cell= worksheet->querySubObject("Cells(int, int)", randomRow, 2);  //获单元格值
                ui->qlb_display_meaning->clear();
                ui->qlb_display_meaning->setText(cell->dynamicCall("Value2()").toString());
            }
        }
    }
    countDown--;
    if(countDown==0)
        countDown=15;
//    qDebug()<<"timer running !";
}

int MainWindow::getExcelRow(){
    // step4: 获取行数，列数
    usedrange = worksheet->querySubObject("UsedRange"); // sheet范围
    QAxObject *rows;
    rows = usedrange->querySubObject("Rows");  // 行
    return rows->property("Count").toInt(); // 行数
}

