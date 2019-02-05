using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{


    /// <summary>
    /// ///////////////////////Export to Excel///////////////////
    /// </summary>

    /*
    private void ExportDatagridToExcel()  //Export to Excel from DataGridView
    {
        int iDGColumns = _dataGridView1ColumnCount();
        int iDGRows = _dataGridView1RowsCount();
        Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
        ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);           //Книга
        ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);    //Таблица.
        ExcelApp.Columns.ColumnWidth = iDGColumns;

        for (int i = 0; i < iDGColumns; i++)
        {
            ExcelApp.Cells[1, 1 + i] = _dataGridView1ColumnHeaderText(i);
            ExcelApp.Columns[1 + i].NumberFormat = "@";
            ExcelApp.Columns[1 + i].AutoFit();
        }

        for (int i = 0; i < iDGRows; i++)
        {
            for (int j = 0; j < iDGColumns; j++)
            { ExcelApp.Cells[i + 2, j + 1] = _dataGridView1CellValue(i, j); }
        }

        ExcelApp.Visible = true;      //Вызываем нашу созданную эксельку.
        ExcelApp.UserControl = true;

        stimerPrev = "";
        _toolStripStatusLabelForeColor(StatusLabel2, Color.Black);
        sLastSelectedElement = "ExportExcel";
        iDGColumns = 0; iDGRows = 0;
        _toolStripStatusLabelSetText(StatusLabel2, "Готово!");
    }
    */
    /*

    // подключить библиотеку Microsoft.Office.Interop.Excel
    // создаем псевдоним для работы с Excel:
    using Excel = Microsoft.Office.Interop.Excel;

        //Объявляем приложение
        Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
        //Отобразить Excel
        ex.Visible = true;
        //Количество листов в рабочей книге
        ex.SheetsInNewWorkbook = 2;
        //Добавить рабочую книгу
        Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
        //Отключить отображение окон с сообщениями
        ex.DisplayAlerts = false;                                       
        //Получаем первый лист документа (счет начинается с 1)
        Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
        //Название листа (вкладки снизу)
        sheet.Name = "Отчет за 13.12.2017";
        //Пример заполнения ячеек
        for (int i = 1; i <= 9; i++)
        {
          for (int j = 1; j < 9; j++)
          sheet.Cells[i, j] = String.Format("Boom {0} {1}", i, j);
        }
        //Захватываем диапазон ячеек
        Excel.Range range1 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 9]);
        //Шрифт для диапазона
        range1.Cells.Font.Name = "Tahoma";
        //Размер шрифта для диапазона
        range1.Cells.Font.Size = 10;
        //Захватываем другой диапазон ячеек
        Excel.Range range2 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 2]);
        range2.Cells.Font.Name = "Times New Roman";
        //Задаем цвет этого диапазона. Необходимо подключить System.Drawing
        range2.Cells.Font.Color = ColorTranslator.ToOle(Color.Green);
        //Фоновый цвет
        range2.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0xFF, 0xFF, 0xCC));

    
    //Расстановка рамок.
   // Расставляем рамки со всех сторон:
        range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
        range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
        range2.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
        range2.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
        range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

    //Цвет рамки можно установить так:
    range2.Borders.Color = ColorTranslator.ToOle(Color.Red);

    //Выравнивания в диапазоне задаются так:
        rangeDate.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        rangeDate.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;


    //Определим задачу: получить сумму диапазона ячеек A4:A10.
    //Для начала снова получим диапазон ячеек:
    Excel.Range formulaRange = sheet.get_Range(sheet.Cells[4, 1], sheet.Cells[9, 1]);

    //Далее получим диапазон вида A4:A10 по адресу ячейки ( [4,1]; [9;1] ) описанному выше:
    string addr = formulaRange.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

    //Теперь в переменной addr у нас хранится строковое значение диапазона ( [4,1]; [9;1] ), то есть A4:A10.
    //Вычисляем формулу:
        //Одна ячейка как диапазон
        Excel.Range r = sheet.Cells[10, 1] as Excel.Range;
        //Оформления
        r.Font.Name = "Times New Roman";
        r.Font.Bold = true;
        r.Font.Color = ColorTranslator.ToOle(Color.Blue);
        //Задаем формулу суммы
        r.Formula = String.Format("=СУММ({0}", addr);

    // Выделение ячейки или диапазона ячеек
    //выделить ячейку или диапазон, как если бы мы выделили их мышкой:
        sheet.get_Range("J3", "J8").Activate();
        //или
        sheet.get_Range("J3", "J8").Select();
        //Можно вписать одну и ту же ячейку, тогда будет выделена одна ячейка.
        sheet.get_Range("J3", "J3").Activate();
        sheet.get_Range("J3", "J3").Select();

        //Чтобы настроить авто ширину и высоту для диапазона, используем такие команды:
        range.EntireColumn.AutoFit(); 
        range.EntireRow.AutoFit();

    //Чтобы получить значение из ячейки, используем такой код:
        //Получение одной ячейки как ранга
        Excel.Range forYach = sheet.Cells[ob + 1, 1] as Excel.Range;
        //Получаем значение из ячейки и преобразуем в строку
        string yach = forYach.Value2.ToString();

    // Чтобы добавить лист и дать ему заголовок, используем следующее:
        var sh = workBook.Sheets;
        Excel.Worksheet sheetPivot = (Excel.Worksheet)sh.Add(Type.Missing, sh[1], Type.Missing, Type.Missing);
        sheetPivot.Name = "Сводная таблица";

   // Добавление разрыва страницы
        //Ячейка, с которой будет разрыв
        Excel.Range razr = sheet.Cells[n, m] as Excel.Range;
        //Добавить горизонтальный разрыв (sheet - текущий лист)
        sheet.HPageBreaks.Add(razr); 
        //VPageBreaks - Добавить вертикальный разрыв

    //Сохраняем документ
        ex.Application.ActiveWorkbook.SaveAs("doc.xlsx", Type.Missing,
          Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
          Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

    //Как открыть существующий документ Excel
        ex.Workbooks.Open(@"C:\Users\Myuser\Documents\Excel.xlsx",
          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
          Type.Missing, Type.Missing);

    //Складываем значения предыдущих 12 ячеек слева
        rang.Formula = "=СУММ(RC[-12]:RC[-1])";

    //Так же во время работы может возникнуть ошибка: метод завершен неверно. Это может означать, что не выбран лист, с которым идет работа.

    //Чтобы выбрать лист, выполните,  где sheetData это нужный лист.
    sheetData.Select(Type.Missing); 

         */
}
