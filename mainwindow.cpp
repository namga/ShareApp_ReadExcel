#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_btnBrowse_clicked()
{
    //open file
    QString fileName = QFileDialog::getOpenFileName(this,tr("Open Spreadsheet"), ".", tr("Spreadsheet files (*.xlsx)"));
    ui->txtLink->setText(fileName);
    if (!fileName.isEmpty())
    {

        int r = QMessageBox::information(this, fileName.toLatin1().mid(fileName.toLatin1().lastIndexOf("/")+1,fileName.toLatin1().length()),
                                         tr(fileName.toLatin1().mid(fileName.toLatin1().lastIndexOf("/")+1,fileName.toLatin1().length())+" "+"has been found./n"
                                                                                                                                             "Do you want to open this file?"),
                                         QMessageBox::Yes | QMessageBox::Default,
                                         QMessageBox::Cancel | QMessageBox::Escape);
        if(r==QMessageBox::Yes)
        {
            // load file into the excel ojbect

            QAxObject* excel = new QAxObject( "Excel.Application", 0 );
            QAxObject* workbooks = excel->querySubObject( "Workbooks" );
            QAxObject* workbook = workbooks->querySubObject( "Open(QString&)", fileName );
            QAxObject* sheets = workbook->querySubObject( "Worksheets" );

            QList<QVariantList> data; //Data list from excel, each QVariantList is worksheet row

            //worksheets count
            int sheetsCount = sheets->dynamicCall("Count()").toInt();
            int columnCount = 0;
            for(int i = 1; i <= sheetsCount ; i++)
            {
                //sheet pointer
                QAxObject* sheet = sheets->querySubObject( "Item( int )", i );

                QAxObject* rows = sheet->querySubObject( "Rows" );
                int rowCount = rows->dynamicCall( "Count()" ).toInt(); //unfortunately, always returns 255, so you have to check somehow validity of cell values

                QAxObject* columns = sheet->querySubObject( "Columns" );
                columnCount = columns->dynamicCall( "Count()" ).toInt(); //similarly, always returns 65535

                //One of possible ways to get column count
                int currentColumnCount = 0;
                for (int col = 1; col < columnCount; col++)
                {
                    QAxObject* cell = sheet->querySubObject( "Cells( int, int )", 1, col );
                    QVariant value = cell->dynamicCall( "Value()" );
                    //qDebug() << "Cell : 1, " << col <<": " << value;

                    if (value.toString().isEmpty())
                        break;
                    else
                        currentColumnCount = col;
                }
                columnCount = currentColumnCount;

                //sheet->dynamicCall( "Calculate()" ); //maybe somewhen it's necessary, but i've found out that cell values are calculated without calling this function. maybe it must be called just to recalculate

                for (int row = 1; row <= rowCount; row++)
                {
                    QVariantList dataRow;
                    bool isEmpty = true; //when all the cells of row are empty, it means that file is at end (of course, it maybe not right for different excel files. it's just criteria to calculate somehow row count for my file)
                    for (int column=1; column <= columnCount; column++)
                    {
                        QAxObject* cell = sheet->querySubObject( "Cells( int, int )", row, column );
                        QVariant value = cell->dynamicCall( "Value()" );
                        if (!value.toString().isEmpty() && isEmpty)
                            isEmpty = false;
                        dataRow.append(value);
                    }
                    if (isEmpty) //criteria to get out of cycle
                        break;
                    data.append(dataRow);
                }

                table_model = new QStandardItemModel(data.size(), columnCount);
                for (int row = 0; row < data.size(); ++row)
                {
                    for (int column = 0; column < columnCount; ++column)
                    {
                        QStandardItem *item = new QStandardItem(data[row].value(column).toString());
                        table_model->setItem(row, column, item);
                    }
                }
            }

            ui->tableView->setModel(table_model);
            ui->tableView->show();

//            tableWidget->setRowCount(rowCount);
//            tableWidget->setColumnCount(columnCount);

//            for(int row = 0; row < rowCount; row++)
//            {
//                for(int column = 0; column < columnCount; column++)
//                {
//                    QTableWidgetItem *item = new QTableWidgetItem(data[0].value(0).toString());
//                    tableWidget->setItem(row, column, item);
//                }
//            }

//            QMessageBox::information(this,"test",data[0].value(0).toString(),QMessageBox::Yes | QMessageBox::Default);

            workbook->dynamicCall("Close()");
            excel->dynamicCall("Quit()");
        }
        else
        {
            this->close();
        }

    }
}
