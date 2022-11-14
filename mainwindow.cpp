#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "xlsxdocument.h"
#include "qfiledialog.h"
#include <QProgressDialog>
using namespace QXlsx;

QString path;
int maxCollum = 512,maxRow = 512;

QStringList sheetNames;
int sheetCount = 0;
QList<QTableWidget*>* tables = new QList<QTableWidget*>();
Document doc;
QPushButton* empty;QPushButton* empty2;
MainWindow::MainWindow(QWidget *parent): QMainWindow(parent), ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->tabWidget->clear();
    connect(ui->save,SIGNAL(released()),this,SLOT(save()));
    connect(ui->open,SIGNAL(released()),this,SLOT(open()));
    connect(ui->newFile,SIGNAL(released()),this,SLOT(createNew()));
    connect(ui->tabWidget,SIGNAL(currentChanged(int)),this,SLOT(addOrRemoveSheet(int)));
    empty = new QPushButton("test");
    empty2 = new QPushButton("test2");
}

MainWindow::~MainWindow()
{
    delete ui;
}
void MainWindow::createNew(){
    ui->tabWidget->clear();
    sheetCount=0;
    sheetNames.clear();
    tables->clear();
    QString num;
    num.setNum(sheetCount);
    sheetCount++;
    QString name = "Лист "+num;
    sheetNames.append(name);
    doc.addSheet(name);
    QTableWidget* newTable = new QTableWidget();
    newTable->setColumnCount(maxCollum);
    newTable->setRowCount(maxRow);
    tables->append(newTable);
    ui->tabWidget->addTab(newTable,name);
    ui->tabWidget->addTab(empty2,"-");
    ui->tabWidget->addTab(empty,"+");
}
void MainWindow::save(){
    if(path.isEmpty()){
        path = QFileDialog::getSaveFileName(this,"Сохранить","","*.xlsx");
    }
    QProgressDialog dialog = QProgressDialog();
    dialog.labelText() = "Saving you file";
    dialog.open();
    for(int i = 0;i<sheetNames.size();i++){
        QString name = sheetNames[i];
        doc.selectSheet(name);
        for(int columm = 0;columm<=maxCollum;columm++){
            for(int row = 0;row<=maxRow;row++){
                int buffRow = row+1;
                int buffCol = columm+1;
                QTableWidgetItem* item = tables->value(i)->item(row,columm);
                if(item != NULL){
                    QVariant value = item->text();
                    doc.write(buffRow,buffCol,value);
                }
            }
        }
    }
    doc.saveAs(path);
    dialog.close();
}
void MainWindow::open(){
    path = QFileDialog::getOpenFileName(0, "Открыть", "", "*.xlsx");
    QProgressDialog dialog = QProgressDialog();
    dialog.labelText() = "Saving you file";
    dialog.open();
    ui->tabWidget->clear();
    QXlsx::Document newdoc(path);
    sheetNames = newdoc.sheetNames();
    if(sheetNames.size()>=1){
        for(int sheetI = 0;sheetI<sheetNames.size();sheetI++){
            QTableWidget* newTable = new QTableWidget();
            newTable->setColumnCount(maxCollum);
            newTable->setRowCount(maxRow);
            newdoc.selectSheet(sheetNames[sheetI]);
            for(int i = 1;i<=maxCollum;i++){
                for(int j = 1;j<=maxRow;j++){
                    Cell* cell = newdoc.cellAt(j,i);
                    if(cell!=NULL){
                    int row = j-1;
                    int col = i-1;
                    QVariant var = cell->readValue();
                    newTable->setItem(row,col,new QTableWidgetItem(var.value<QString>()));
                    }
                }
            }
            addNewSheet(sheetI,newTable,sheetNames[sheetI]);
        }
        ui->tabWidget->addTab(empty,"-");
        ui->tabWidget->addTab(empty2,"+");
    }
    dialog.close();
}
void MainWindow::addOrRemoveSheet(int i){
    if(!tables->empty()){
    if(i==tables->count() && sheetNames.count()>1){
        doc.deleteSheet(sheetNames[i-1]);
        sheetCount--;
        sheetNames.removeAt(i-1);
        tables->removeAt(i-1);
        ui->tabWidget->removeTab(i-1);
        ui->tabWidget->setCurrentIndex(i-1);
    }
    else if(i == tables->count()+1){
        QString num;
        num.setNum(sheetCount);
        sheetCount++;
        QString name = "Лист "+num;
        sheetNames.append(name);
        doc.addSheet(name);
        QTableWidget* newTable = new QTableWidget();
        newTable->setColumnCount(maxCollum);
        newTable->setRowCount(maxRow);
        tables->append(newTable);
        ui->tabWidget->insertTab(i-1,newTable,name);
        ui->tabWidget->setCurrentIndex(i-1);
    }
    else if(i==tables->count() || i == tables->count()+1){
        ui->tabWidget->setCurrentIndex(sheetNames.indexOf(sheetNames.last()));
    }
    }
}
void MainWindow::addNewSheet(int i,QTableWidget* newTable,QString name){
        sheetCount++;
        doc.addSheet(name);
        tables->append(newTable);
        ui->tabWidget->insertTab(i,newTable,name);
        ui->tabWidget->setCurrentIndex(i);
}
